VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTTickfileReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ITickfileReader

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const MaxTickTableIndex As Long = 9999
Private Const OneMinute As Double = 1 / 1440
Private Const TimeoutSecs As Long = 10

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type FieldDescriptor
    Name        As String
    DataType    As Long
    Size        As Long
End Type

Private Type Tick
    TickType    As TickTypes
    timestamp   As Date
    Price       As Double
    Size        As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileInputServiceConsumer
Private mDataConsumer As IStreamingDataConsumer

Private mProviderKey As String

Private WithEvents mConnectionTimer As TimerUtils.IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1
Private WithEvents mTimeoutTimer As TimerUtils.IntervalTimer
Attribute mTimeoutTimer.VB_VarHelpID = -1
Private mElapsedTimer As TimerUtils.ElapsedTimer

Private mfWinsock As fWinsock
Private WithEvents mSocket As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mConnectionRetryIntervalSecs As Long
Private mLoggedOn As Boolean
Private mPassword As String
Private mTotalBytesInBuffer As Long
Private mTotalBytesProcessed As Long ' number of bytes processed in the current socket buffer

Private mFieldTable() As FieldDescriptor

Private mTickTable(MaxTickTableIndex) As Tick
Private mNumUnplayedTicks As Long
Private mNumPlayedTicks As Long
Private mNextTickID As Long
Private mPrevReplayTickID As Long

Private mAllTicksReceived As Boolean
Private mTickRequestPending As Boolean

Private mReplayProgressEventFrequency As Long

Private mInputTickfileSpecifier As TradeBuildSP.TickfileSpec
'Private mContract As TradeBuildSP.IContract

Private mMonths(12) As String

Private mStarttimeLocal As Date
Private mEndtimeLocal As Date

Private mTimestampOffset As Date    ' used to convert timestamps received from QT
                                    ' (which are in Eastern Standard Time) to
                                    ' local time

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

mMonths(1) = "Jan"
mMonths(2) = "Feb"
mMonths(3) = "Mar"
mMonths(4) = "Apr"
mMonths(5) = "May"
mMonths(6) = "Jun"
mMonths(7) = "Jul"
mMonths(8) = "Aug"
mMonths(9) = "Sep"
mMonths(10) = "Oct"
mMonths(11) = "Nov"
mMonths(12) = "Dec"

Set mTimeoutTimer = New TimerUtils.IntervalTimer
mTimeoutTimer.TimerIntervalSecs = TimeoutSecs
mTimeoutTimer.RepeatNotifications = False

Set mElapsedTimer = New TimerUtils.ElapsedTimer

mPrevReplayTickID = -1
End Sub

'================================================================================
' ITickfileReader Interface Members
'================================================================================

Private Sub ITickfileReader_CloseInputFile()
ReleaseSocket
If Not mConnectionTimer Is Nothing Then
    mConnectionTimer.StopTimer
    Set mConnectionTimer = Nothing
End If
End Sub

Private Property Let ITickfileReader_Contract(ByVal RHS As TradeBuildSP.IContract)
'Set mContract = RHS
'If mLoggedOn Then RequestHistoricTickData
End Property

Private Property Get ITickfileReader_Contract() As TradeBuildSP.IContract
Set ITickfileReader_Contract = mInputTickfileSpecifier.Contract
End Property

Private Property Get ITickfileReader_ContractSpecifier() As TradeBuildSP.IContractSpecifier
Set ITickfileReader_ContractSpecifier = mInputTickfileSpecifier.Contract.Specifier
End Property

Private Sub ITickfileReader_FireNextTick()
FireNextTick
End Sub

Private Property Get ITickfileReader_FirstTickFileTime() As Date
ITickfileReader_FirstTickFileTime = FirstTickFileTime
End Property

Private Property Let ITickfileReader_InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpec)
InputTickfileSpecifier = RHS
End Property

Private Property Get ITickfileReader_InputTickfileSpecifier() As TradeBuildSP.TickfileSpec
ITickfileReader_InputTickfileSpecifier = InputTickfileSpecifier
End Property

Private Property Let ITickfileReader_ReplayProgressEventFrequency(ByVal RHS As Long)
If RHS = 0 Then
    mReplayProgressEventFrequency = 1
Else
    mReplayProgressEventFrequency = RHS
End If
End Property

Private Property Get ITickfileReader_ReplayProgressEventFrequency() As Long
ITickfileReader_ReplayProgressEventFrequency = mReplayProgressEventFrequency
End Property

Private Function ITickfileReader_Supports( _
                            ByVal Capabilities As Long, _
                            Optional ByVal FormatIdentifier As String) As Boolean
ITickfileReader_Supports = gSupports(Capabilities, FormatIdentifier)
End Function

Private Property Get ITickfileReader_tickfileSizeBytes() As Long
ITickfileReader_tickfileSizeBytes = 0
End Property

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
ConnectHelper
End Sub

'================================================================================
' mTimeoutTimer Event Handlers
'================================================================================

Private Sub mTimeoutTimer_TimerExpired()
ReleaseSocket
mServiceConsumer.error StandardSPErrorCodes.TFDataSourceNotResponding, _
                        "QuoteTracker Server is not responding"
End Sub

'================================================================================
' mSocket Event Handlers
'================================================================================

Private Sub mSocket_Connect()
RequestFieldDescription
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)

On Error GoTo err

If Not mTimeoutTimer Is Nothing Then mTimeoutTimer.StopTimer

mTotalBytesInBuffer = bytesTotal
mTotalBytesProcessed = 0

Debug.Print "Socket data size = " & mTotalBytesInBuffer
ProcessSocketData

If mTotalBytesProcessed >= mTotalBytesInBuffer And _
    Not mAllTicksReceived _
Then
    mTimeoutTimer.StartTimer
End If

If mTickRequestPending Then FireNextTick
Exit Sub

err:
If err.Number <> ErrorCodes.LongJump Then
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.UnexpectedError, _
        "Unexpected error: " & _
        err.Number & ": " & err.Description, _
        gName
End If
ReleaseSocket
mServiceConsumer.TickfileReplayComplete
End Sub

Private Sub mSocket_Error(ByVal Number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal Source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Dim errorNum As MSWinsockLib.ErrorConstants
Dim timestamp As Date: timestamp = mCommonServiceConsumer.GetTimestamp

errorNum = Number
Select Case errorNum
Case sckAddressNotAvailable, _
    sckNetworkSubsystemFailed, _
    sckNetworkUnreachable, _
    sckNetReset, _
    sckConnectAborted, _
    sckConnectionRefused, _
    sckHostNotFound, _
    sckHostNotFoundTryAgain
    If Not mSocket.State = sckConnected Then
        ReleaseSocket
        If mConnectionRetryIntervalSecs <> 0 Then
            mServiceConsumer.error StandardSPErrorCodes.TFRetryConnectDataSource, _
                                        "Retrying connection to QuoteTracker"
            If mConnectionTimer Is Nothing Then
                Set mConnectionTimer = New IntervalTimer
                mConnectionTimer.RepeatNotifications = False
                mConnectionTimer.TimerIntervalSecs = mConnectionRetryIntervalSecs
            End If
            mConnectionTimer.StartTimer
        Else
            mServiceConsumer.error StandardSPErrorCodes.TFCantConnectDataSource, _
                                        "Can't connect to QuoteTracker"
                                        
        End If
    Else
        mCommonServiceConsumer.ServiceProviderError errorNum, _
                                    Description, _
                                    gName
    End If
Case Else
    mCommonServiceConsumer.ServiceProviderError errorNum, _
                                Description, _
                                gName
    If mTickRequestPending Then
        mServiceConsumer.TickfileReplayComplete
        mAllTicksReceived = True
        mTickRequestPending = False
        ReleaseSocket
    End If
End Select

End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let commonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property


Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Get FirstTickFileTime() As Date
' not supported
End Property

Friend Property Let InputTickfileSpecifier(RHS As TradeBuildSP.TickfileSpec)
mInputTickfileSpecifier = RHS
If mLoggedOn Then
    RequestHistoricTickData
End If
End Property

Friend Property Get InputTickfileSpecifier() As TradeBuildSP.TickfileSpec
InputTickfileSpecifier = mInputTickfileSpecifier
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let StreamingDataConsumer(ByVal RHS As IStreamingDataConsumer)
Set mDataConsumer = RHS
End Property

Friend Property Let TickfileInputServiceConsumer(ByVal RHS As ITickfileInputServiceConsumer)
Set mServiceConsumer = RHS
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub Connect(ByVal server As String, _
                            ByVal port As Long, _
                            ByVal password As String)
mPassword = password
mServer = server
mPort = port
ConnectHelper
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addTickToTable(ByRef pTick As Tick)
mTickTable(mNextTickID) = pTick
mNextTickID = (mNextTickID + 1) Mod MaxTickTableIndex
mNumUnplayedTicks = mNumUnplayedTicks + 1
End Sub

Private Sub ConnectHelper()
Set mfWinsock = New fWinsock
Set mSocket = mfWinsock.Winsock1
mSocket.Connect mServer, mPort
End Sub

Private Sub FireNextTick()

If mNumUnplayedTicks = 0 Then
    If mAllTicksReceived Then
        mServiceConsumer.TickfileReplayComplete
        Exit Sub
    Else
        ProcessSocketData
        If mNumUnplayedTicks = 0 Then
            If mAllTicksReceived Then
                mServiceConsumer.TickfileReplayComplete
                Exit Sub
            Else
                mTickRequestPending = True
                Exit Sub
            End If
        End If
    End If
End If

mTickRequestPending = False

mPrevReplayTickID = (mPrevReplayTickID + 1) Mod MaxTickTableIndex

With mTickTable(mPrevReplayTickID)
    mNumUnplayedTicks = mNumUnplayedTicks - 1
    mNumPlayedTicks = mNumPlayedTicks + 1
    
    If mPrevReplayTickID Mod mReplayProgressEventFrequency = 0 Then
        mServiceConsumer.ReplayProgress .timestamp, _
                                mNumPlayedTicks, _
                                0#
    End If

    Select Case .TickType
    Case TickTypes.Bid
        mDataConsumer.Bid .timestamp, .Price, .Size
    Case TickTypes.Ask
        mDataConsumer.Ask .timestamp, .Price, .Size
    Case TickTypes.Last
        mDataConsumer.Trade .timestamp, .Price, .Size
    Case TickTypes.Volume
        mDataConsumer.Volume .timestamp, .Size
    Case TickTypes.PrevClose
        mDataConsumer.PreviousClose .timestamp, .Price
    Case TickTypes.High
        mDataConsumer.High .timestamp, .Price
    Case TickTypes.Low
        mDataConsumer.Low .timestamp, .Price
    Case TickTypes.OpenInterest
        mDataConsumer.OpenInterest .timestamp, .Size
    End Select
End With

End Sub

Private Function GetFieldValue( _
                            ByVal fieldID As Long, _
                            ByRef bytesProcessed As Long) As Variant
Dim value As Variant
Dim padding() As Byte

With mFieldTable(fieldID)
    Select Case .DataType
    Case vbString
        If .Size <> 0 Then
            mSocket.GetData value, vbString, .Size
            bytesProcessed = bytesProcessed + .Size
        Else
            Dim letter As Byte
            Dim valueString As String
            Do
                mSocket.GetData letter, vbByte, 1
                bytesProcessed = bytesProcessed + 1
                If letter > 0 Then
                    valueString = valueString + Chr(letter)
                End If
            Loop While letter > 0
            value = valueString
        End If
    Case vbLong
        mSocket.GetData value, vbLong
        bytesProcessed = bytesProcessed + 4
        If .Size = 8 Then
            mSocket.GetData padding, vbByte + vbArray, 4
            bytesProcessed = bytesProcessed + 4
        End If
    Case vbDouble
        mSocket.GetData value, vbDouble
        bytesProcessed = bytesProcessed + 8
    Case vbByte
        If .Size = 1 Then
            mSocket.GetData value, vbByte, 1
            bytesProcessed = bytesProcessed + 1
        Else
            mSocket.GetData value, vbByte + vbArray, .Size
            bytesProcessed = bytesProcessed + .Size
        End If
    End Select
End With
GetFieldValue = value
End Function

Private Function GotDataForRecord( _
                            ByRef recordID As String, _
                            ByRef recordLength As Integer, _
                            ByRef bytesProcessed As Long _
                            ) As Boolean
Static expectingMoreData As Boolean
If Not expectingMoreData Then
    bytesProcessed = 0
    If (mTotalBytesInBuffer - mTotalBytesProcessed) >= 4 Then
        mSocket.GetData recordID, vbString, 2
        mSocket.GetData recordLength, vbInteger, 2
        bytesProcessed = 4
        If (mTotalBytesInBuffer - mTotalBytesProcessed) < recordLength Then
            expectingMoreData = True
            GotDataForRecord = False
            Exit Function
        End If
    Else
        expectingMoreData = True
        GotDataForRecord = False
        Exit Function
    End If
Else
    ' see if we've now got enough data for the record
    If (mTotalBytesInBuffer + bytesProcessed) < recordLength Then
        Exit Function
    End If
    expectingMoreData = False
End If
54 GotDataForRecord = True
End Function

Private Sub ProcessError( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim XMLdoc As DOMDocument30
Dim XMLData As String
Dim error As IXMLDOMElement
Dim errorCode As IXMLDOMElement
Dim errorDesc As IXMLDOMElement

mSocket.GetData XMLData, vbString, recordLength - bytesProcessed
Debug.Assert Len(XMLData) = recordLength - bytesProcessed
bytesProcessed = bytesProcessed + Len(XMLData)

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML XMLData
If XMLdoc.parseError.errorCode <> 0 Then
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.ErrorCannotBeParsed, _
            "Error XML cannot be parsed", _
            gName
    err.Raise ErrorCodes.LongJump
End If

Set error = XMLdoc.selectSingleNode("ERROR")
Set errorCode = error.selectSingleNode("ERRORCODE")
Set errorDesc = error.selectSingleNode("ERRORDESC")

mCommonServiceConsumer.ServiceProviderError SPErrorCodes.ErrorFromQuoteTracker, _
        "QT API Error " & errorCode.Text & ": " & errorDesc.Text, _
        gName
err.Raise ErrorCodes.LongJump
End Sub

Private Sub ProcessFieldDescriptions( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim XMLdoc As DOMDocument30
Dim XMLData As String
Dim XMLNodeList As IXMLDOMNodeList
Dim fields As IXMLDOMElement
Dim field As IXMLDOMElement
Dim fieldID As FieldTypes

mSocket.GetData XMLData, vbString, recordLength - bytesProcessed
Debug.Assert Len(XMLData) = recordLength - bytesProcessed
bytesProcessed = bytesProcessed + Len(XMLData)

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML XMLData
If XMLdoc.parseError.errorCode <> 0 Then
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.FieldDescriptionCannotBeParsed, _
            "Field descriptions XML cannot be parsed", _
            gName
    err.Raise ErrorCodes.LongJump
End If

Set fields = XMLdoc.selectSingleNode("FIELDS")
Set XMLNodeList = fields.selectNodes("FIELD")

ReDim mFieldTable(9) As FieldDescriptor

For Each field In XMLNodeList
    fieldID = field.getAttribute("ID")
    If fieldID > UBound(mFieldTable) Then
        ReDim Preserve mFieldTable(fieldID + 9) As FieldDescriptor
    End If
    With mFieldTable(fieldID)
        Dim typeAtt As String
        .Name = field.Text
        .Size = CLng("0" & field.getAttribute("Size"))
        typeAtt = LCase(field.getAttribute("Type"))
        Select Case typeAtt
        Case "text"
            .DataType = vbString
        Case "char"
            .DataType = vbByte
        Case "int"
            .DataType = vbLong
        Case "float"
            .DataType = vbDouble
        End Select
    End With
Next

RequestLogin

End Sub

Private Sub ProcessedHistoricalDataEnd( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim fieldIdRaw As Integer
Dim fieldID As FieldTypes

mAllTicksReceived = True

Do
    mSocket.GetData fieldIdRaw, vbInteger, 2
    fieldID = CLng(fieldIdRaw)
    bytesProcessed = bytesProcessed + 2
    GetFieldValue fieldID, bytesProcessed
Loop While bytesProcessed < recordLength

ReleaseSocket
End Sub

Private Sub ProcessHistoricalTick( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim fieldIdRaw As Integer
Dim fieldID As FieldTypes

Dim timestamp As Date

Static firstInScopeTickProcessed As Boolean
Static prevBid As Double
Static prevAsk As Double
Static prevVolume As Long

Dim thisBid As Double
Dim thisAsk As Double
Dim thisLast As Double
Dim thisVolume As Long

Dim bidTick As Tick
Dim askTick As Tick
Dim lastTick As Tick
Dim volumeTick As Tick

Do
    mSocket.GetData fieldIdRaw, vbInteger, 2
    fieldID = CLng(fieldIdRaw)
    bytesProcessed = bytesProcessed + 2
    
    Select Case fieldID
    Case FieldTypes.BidPrice
        thisBid = GetFieldValue(fieldID, bytesProcessed)
        If thisBid <> prevBid Then
            bidTick.TickType = TickTypes.Bid
            bidTick.Price = thisBid
            prevBid = thisBid
        End If
    Case FieldTypes.AskPrice
        thisAsk = GetFieldValue(fieldID, bytesProcessed)
        If thisAsk <> prevAsk Then
            askTick.TickType = TickTypes.Ask
            askTick.Price = thisAsk
            prevAsk = thisAsk
        End If
    Case FieldTypes.LastPrice
        thisLast = GetFieldValue(fieldID, bytesProcessed)
    Case FieldTypes.Volume
        thisVolume = GetFieldValue(fieldID, bytesProcessed)
        If thisVolume <> prevVolume Then
            If thisVolume > prevVolume Then
                lastTick.TickType = TickTypes.Last
                lastTick.Price = thisLast
                lastTick.Size = thisVolume - prevVolume
            End If
        
            volumeTick.TickType = TickTypes.Volume
            volumeTick.Size = thisVolume
            prevVolume = thisVolume
        End If
    Case FieldTypes.timestamp
        Dim timestampString As String
        timestampString = GetFieldValue(fieldID, bytesProcessed)
        timestamp = CDate(mMonths(Mid$(timestampString, 3, 2)) & " " & _
                            Left$(timestampString, 2) & " " & _
                            Mid$(timestampString, 5, 4) & " " & _
                            Mid$(timestampString, 9, 2) & ":" & _
                            Mid$(timestampString, 11, 2) & ":" & _
                            Mid(timestampString, 13, 2)) + _
                            mTimestampOffset
                            
    Case Else
        GetFieldValue fieldID, bytesProcessed
    End Select

Loop While bytesProcessed < recordLength

If timestamp >= mEndtimeLocal Then Stop

If timestamp >= mStarttimeLocal And _
    timestamp < mEndtimeLocal _
Then
    If bidTick.TickType <> TickTypes.None Then
        bidTick.timestamp = timestamp
        addTickToTable bidTick
    ElseIf Not firstInScopeTickProcessed Then
        If prevBid <> 0# Then
            bidTick.TickType = TickTypes.Bid
            bidTick.Price = prevBid
            bidTick.timestamp = timestamp
            addTickToTable bidTick
        End If
    End If
    If askTick.TickType <> TickTypes.None Then
        askTick.timestamp = timestamp
        addTickToTable askTick
    ElseIf Not firstInScopeTickProcessed Then
        firstInScopeTickProcessed = True
        If prevAsk <> 0# Then
            askTick.TickType = TickTypes.Ask
            askTick.Price = prevAsk
            askTick.timestamp = timestamp
            addTickToTable askTick
        End If
    End If
    If lastTick.TickType <> TickTypes.None Then
        lastTick.timestamp = timestamp
        addTickToTable lastTick
    End If
    If volumeTick.TickType <> TickTypes.None Then
        volumeTick.timestamp = timestamp
        addTickToTable volumeTick
    End If
    
    firstInScopeTickProcessed = True
End If

End Sub

Private Sub ProcessLoginResponse( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim XMLdoc As DOMDocument30
Dim XMLData As String
Dim login As IXMLDOMElement
Dim status As IXMLDOMElement
Dim statusCode As IXMLDOMElement
Dim statusDesc As IXMLDOMElement

mSocket.GetData XMLData, vbString, recordLength - bytesProcessed
Debug.Assert Len(XMLData) = recordLength - bytesProcessed
bytesProcessed = bytesProcessed + Len(XMLData)

Set XMLdoc = New DOMDocument30
XMLdoc.loadXML XMLData
If XMLdoc.parseError.errorCode <> 0 Then
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.LogonResponseCannotBeParsed, _
            "Logon response XML cannot be parsed", _
            gName
    err.Raise ErrorCodes.LongJump
End If

Set login = XMLdoc.selectSingleNode("LOGIN")
Set statusCode = login.selectSingleNode("STATUSCODE")
Select Case statusCode.Text
Case "0"
    ' login successful
    mLoggedOn = True
    If Not mInputTickfileSpecifier.Contract Is Nothing Then
        RequestHistoricTickData
    End If
    mServiceConsumer.Ready mInputTickfileSpecifier, 0
Case "1"
    ' login failed - bad password
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.PasswordInvalid, _
            "Logon password invalid", _
            gName
    mServiceConsumer.error StandardSPErrorCodes.TFCantConnectDataSource, _
                                    "Can't connect to QuoteTracker"
    err.Raise ErrorCodes.LongJump
Case Else
    Set statusDesc = XMLdoc.selectSingleNode("STATUSDESC")
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.UnexpectedError, _
            statusCode.Text & ": " & statusDesc.Text, _
            gName
    mServiceConsumer.error StandardSPErrorCodes.TFCantConnectDataSource, _
                                    "Can't connect to QuoteTracker"
    err.Raise ErrorCodes.LongJump
End Select
End Sub

Private Function ProcessRecord() As Boolean
Static recordID As String
Static recordLength As Integer
Static bytesProcessed As Long  ' number of bytes processed in the current record

If TickTableFull Then Exit Function

If Not GotDataForRecord(recordID, recordLength, bytesProcessed) Then Exit Function

Select Case recordID
Case "DE"       ' field descriptions
    ProcessFieldDescriptions recordLength, bytesProcessed
Case "ER"       ' error
    ProcessError recordLength, bytesProcessed
Case "LO"       ' login response
    ProcessLoginResponse recordLength, bytesProcessed
Case "HT"       ' historical tick
    ProcessHistoricalTick recordLength, bytesProcessed
Case "HE"       ' end of historical data
    ProcessedHistoricalDataEnd recordLength, bytesProcessed
Case Else
    Dim data As String
    ' get the rest of the message off the socket
    mSocket.GetData data, vbString, recordLength - bytesProcessed
    bytesProcessed = bytesProcessed + Len(data)
End Select

mTotalBytesProcessed = mTotalBytesProcessed + bytesProcessed
ProcessRecord = True
End Function

Private Sub ProcessSocketData()
If mTotalBytesProcessed >= mTotalBytesInBuffer Then Exit Sub

'If mNumUnplayedTicks < 10 Then
    Debug.Print "Tick queue length=" & mNumUnplayedTicks
    
    mElapsedTimer.StartTiming
    
    Do While ProcessRecord And _
            mTotalBytesProcessed < mTotalBytesInBuffer And _
            mElapsedTimer.ElapsedTimeMicroseconds < 20000!
    Loop

    Debug.Print "Tick queue length=" & mNumUnplayedTicks
    
'End If
End Sub

Private Sub ReleaseSocket()
If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
    Unload mfWinsock
    Set mfWinsock = Nothing
End If
End Sub

Private Sub RequestFieldDescription()
Dim qtRecDesc As String

qtRecDesc = "<RECDESC>" & _
            "<ReqType>SNAP</ReqType>" & _
            "</RECDESC>" & _
            Chr(255)

mSocket.SendData qtRecDesc
End Sub

Private Sub RequestHistoricTickData()
Dim qtHist As String
Dim starttimeEST As Date
Dim endtimeEST As Date

If mInputTickfileSpecifier.EntireSession Then
    mInputTickfileSpecifier.Contract.GetSessionTimes mInputTickfileSpecifier.From, mStarttimeLocal, mEndtimeLocal
Else
    mStarttimeLocal = mInputTickfileSpecifier.From
    mEndtimeLocal = mInputTickfileSpecifier.To
End If
starttimeEST = gConvertLocalTimeToEST(mStarttimeLocal)
endtimeEST = gConvertLocalTimeToEST(mEndtimeLocal)
mTimestampOffset = mStarttimeLocal - starttimeEST

' NB - we ask for the data to start one minute earlier than requested, because
' otherwise the first trade gets the size set to volume (too bad if there hasn't been
' a trade in the previous minute!!)

qtHist = "<GetHistory>" & _
        "<ReqType>TICK</ReqType>" & _
        "<Symbol>" & _
        mInputTickfileSpecifier.Contract.ProviderID(mProviderKey) & _
        "</Symbol>" & _
        "<FROM>" & _
        Format(starttimeEST - OneMinute, "yyyy/mm/dd hh:mm:ss") & _
        "</FROM>" & _
        "<TO>" & _
        Format(endtimeEST, "yyyy/mm/dd hh:mm:ss") & _
        "</TO>" & _
        "</GetHistory>" & _
        Chr(255)

mSocket.SendData qtHist
mTimeoutTimer.StartTimer

End Sub

Private Sub RequestLogin()
Dim qtLogin As String
qtLogin = "<LOGIN><PASSWORD>mPassword</PASSWORD>" & _
            "<AppName>TBQTSP</AppName>" & _
            "<AppVer>1</AppVer>" & _
            "</LOGIN>" + _
            Chr(255)

mSocket.SendData qtLogin
End Sub

Private Function TickTableFull() As Boolean
' check if there is room for the possible 4 ticks that could be
' created from this record
If ((mNextTickID > mPrevReplayTickID) And _
    ((mNextTickID + 4) > (mPrevReplayTickID + MaxTickTableIndex))) _
    Or _
    ((mNextTickID <= mPrevReplayTickID) And _
    ((mNextTickID + 4) > mPrevReplayTickID)) _
Then
    TickTableFull = True
End If
End Function

