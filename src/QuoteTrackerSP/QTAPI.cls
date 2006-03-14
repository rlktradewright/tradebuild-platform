VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QTAPI"
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

'================================================================================
' Events
'================================================================================

Event Connected()
    ' Raised when a successful connection to QuoteTracker has been achieved.

Event ConnectFailed(ByVal Description As String, _
                ByVal retrying As Boolean)
    ' Raised when connection to QuoteTracker has failed and the connection
    ' retry interval is zero.

Event Connecting()
    ' Raised when an attempt to connect to QuoteTracker is initiated.

Event ConnectionClosed( _
                ByVal reconnecting As Boolean)
    ' Raised when the connection to QuoteTracker is disconnected.

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type FieldDescriptor
    Name        As String
    DataType    As Long
    size        As Long
End Type

Private Type HistoricalTickDataRequestEntry
    inUse               As Boolean
    cancelling          As Boolean
    symbol              As String
    requester           As QTTickfileReader
    firstInScopeTickProcessed As Boolean
    prevBid             As Double
    prevAsk             As Double
    prevVolume          As Long
    startTimeLocal      As Date
    endTimeLocal        As Date
    timestampOffset     As Date ' used to convert timestamps received from QT
                                ' (which are in Eastern Standard Time) to
                                ' local time
End Type

Private Type L1QuotesRequestEntry
    inUse               As Boolean
    cancelling          As Boolean
    reuseTime           As Date
    symbol              As String
    requester           As QTRealtimeDataReader
    timestampOffset     As Date ' used to convert timestamps received from QT
                                ' (which are in Eastern Standard Time) to
                                ' local time
    bid                 As Double
    bidSize             As Long
    ask                 As Double
    askSize             As Long
    last                As Double
    lastSize            As Long
    high                As Double
    low                 As Double
    prevClose           As Double
    volume              As Long
    openInterest        As Long
    timestamp           As Date
End Type

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer

Private WithEvents mConnectionTimer As TimerUtils.IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1

Private WithEvents mSocket As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mPort As Long
Private mServer As String
Private mConnectionRetryIntervalSecs As Long
Private mConnectionState As ConnectionStates
Private mDisconnectedByApplication As Boolean

Private mProviderKey As String

Private mPassword As String

Private mLogLevel As LogLevels

Private mMonths(12) As String

Private mTotalBytesInBuffer As Long
Private mTotalBytesProcessed As Long ' number of bytes processed in the current socket buffer

Private mFieldTable() As FieldDescriptor

Private mHistoricalTickDataRequestTable() As HistoricalTickDataRequestEntry

Private mL1QuotesRequestTable() As L1QuotesRequestEntry

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

ReDim mHistoricalTickDataRequestTable(5) As HistoricalTickDataRequestEntry

ReDim mL1QuotesRequestTable(10) As L1QuotesRequestEntry

End Sub

Private Sub Class_Terminate()
Debug.Print "QTSP.QTAPI terminated"
End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
connectToQT
End Sub

'================================================================================
' mTimeoutTimer Event Handlers
'================================================================================

Private Sub mTimeoutTimer_TimerExpired()
ReleaseSocket
'mServiceConsumer.error StandardSPErrorCodes.TFDataSourceNotResponding, _
'                        "QuoteTracker Server is not responding"
End Sub

'================================================================================
' mSocket Event Handlers
'================================================================================

Private Sub mSocket_Close()
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
If mDisconnectedByApplication Then
    RaiseEvent ConnectionClosed(False)
    logMessage "Connection to QT closed: " & calcConnectionString, LogLevelLow
Else
    retryConnection
    RaiseEvent ConnectionClosed( _
                            IIf(mConnectionRetryIntervalSecs = 0, False, True))
    logMessage "Connection to QT closed" & _
                IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                calcConnectionString, _
                LogLevelLow
End If
End Sub

Private Sub mSocket_Connect()
mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting
logMessage "Connecting to QT: " & calcConnectionString, LogLevelLow
RequestFieldDescription
End Sub

Private Sub mSocket_DataArrival(ByVal bytesTotal As Long)

On Error GoTo err

mTotalBytesInBuffer = bytesTotal
mTotalBytesProcessed = 0

Debug.Print "Socket data received: size = " & mTotalBytesInBuffer
ProcessSocketData

Exit Sub

err:
ReleaseSocket
If err.Number <> LocalErrorCodes.LongJump Then
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.UnexpectedError, _
        "Unexpected error: " & _
        err.Number & ": " & err.Description, _
        App.ProductName
End If
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
    If mConnectionState = ConnectionStates.ConnNotConnected Then
        ReleaseSocket
        logMessage "Failed to connect to QuoteTracker" & _
                    IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                    Description & ": " & _
                    calcConnectionString, _
                    LogLevelLow
        RaiseEvent ConnectFailed(Description, _
                                IIf(mConnectionRetryIntervalSecs = 0, False, True))
        retryConnection
    Else
        mConnectionState = ConnectionStates.ConnNotConnected
        ReleaseSocket
        retryConnection
        logMessage "Connection to QuoteTracker closed: " & _
                    IIf(mConnectionRetryIntervalSecs = 0, "", " (retrying)") & ": " & _
                    Description & ": " & _
                    calcConnectionString, _
                    LogLevelLow
        RaiseEvent ConnectionClosed(IIf(mConnectionRetryIntervalSecs = 0, False, True))
    End If
Case Else
    logMessage "RuntimeException in QTSP.QTAPI::mSocket_error" & vbCrLf & _
                "Socket error:" & vbCrLf & _
                "Number: " & Number & vbCrLf & _
                "Description: " & Description & vbCrLf & _
                "Scode: " & Scode, _
                LogLevelLow
    err.Raise ErrorCodes.RuntimeException, _
                "QTSP.QTAPI::mSocket_error", _
                "Socket error:" & vbCrLf & _
                "Number: " & Number & vbCrLf & _
                "Description: " & Description & vbCrLf & _
                "Scode: " & Scode
End Select

End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let CommonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property


Friend Property Let ConnectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Get connectionState() As ConnectionStates
connectionState = mConnectionState
End Property

Friend Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Friend Property Let password(ByVal value As String)
mPassword = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub cancelHistoricTickData( _
                ByVal requestId As Long)
If mHistoricalTickDataRequestTable(requestId).inUse Then
    mHistoricalTickDataRequestTable(requestId).cancelling = True
End If
End Sub

Friend Sub Connect()
connectToQT
End Sub

Friend Sub disconnect()
mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected
ReleaseSocket
RaiseEvent ConnectionClosed(False)
End Sub

Friend Function RequestHistoricTickData( _
                ByRef tickfileSpecifier As TradeBuildSP.TickfileSpec, _
                ByVal requester As QTTickfileReader) As Long
Dim qtHist As String
Dim starttimeEST As Date
Dim endtimeEST As Date
Dim i As Long

For i = 0 To UBound(mHistoricalTickDataRequestTable)
    If Not mHistoricalTickDataRequestTable(i).inUse Then Exit For
Next

If i > UBound(mHistoricalTickDataRequestTable) Then
    ReDim Preserve mHistoricalTickDataRequestTable(UBound(mHistoricalTickDataRequestTable) + 5) As HistoricalTickDataRequestEntry
End If

mHistoricalTickDataRequestTable(i).inUse = True
Set mHistoricalTickDataRequestTable(i).requester = requester
mHistoricalTickDataRequestTable(i).symbol = UCase$(tickfileSpecifier.contract.ProviderID(mProviderKey))

If tickfileSpecifier.EntireSession Then
    tickfileSpecifier.contract.GetSessionTimes tickfileSpecifier.From, mHistoricalTickDataRequestTable(i).startTimeLocal, mHistoricalTickDataRequestTable(i).endTimeLocal
Else
    mHistoricalTickDataRequestTable(i).startTimeLocal = tickfileSpecifier.From
    mHistoricalTickDataRequestTable(i).endTimeLocal = tickfileSpecifier.To
End If
starttimeEST = gConvertLocalTimeToEST(mHistoricalTickDataRequestTable(i).startTimeLocal)
endtimeEST = gConvertLocalTimeToEST(mHistoricalTickDataRequestTable(i).endTimeLocal)
mHistoricalTickDataRequestTable(i).timestampOffset = mHistoricalTickDataRequestTable(i).startTimeLocal - starttimeEST

' NB - we ask for the data to start one minute earlier than requested, because
' otherwise the first trade gets the size set to volume (too bad if there hasn't been
' a trade in the previous minute!!)

qtHist = "<GetHistory>" & _
        "<ReqType>TICK</ReqType>" & _
        "<Symbol>" & _
        UCase$(tickfileSpecifier.contract.ProviderID(mProviderKey)) & _
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

End Function

Friend Function subscribeToL1Quotes( _
                ByVal pContract As TradeBuildSP.IContract, _
                ByVal requester As QTRealtimeDataReader) As Long
Dim qtQuote As String
Dim i As Long

For i = 0 To UBound(mL1QuotesRequestTable)
    If Not mL1QuotesRequestTable(i).inUse Or _
        (mL1QuotesRequestTable(i).cancelling And mL1QuotesRequestTable(i).reuseTime <= Now) _
    Then Exit For
Next

If i > UBound(mL1QuotesRequestTable) Then
    ReDim Preserve mL1QuotesRequestTable(UBound(mL1QuotesRequestTable) + 10) As L1QuotesRequestEntry
End If

mL1QuotesRequestTable(i).inUse = True
mL1QuotesRequestTable(i).cancelling = False
mL1QuotesRequestTable(i).reuseTime = 0
Set mL1QuotesRequestTable(i).requester = requester
mL1QuotesRequestTable(i).symbol = UCase$(pContract.ProviderID(mProviderKey))

mL1QuotesRequestTable(i).timestampOffset = Now - gConvertLocalTimeToEST(Now)

qtQuote = "<L1Quotes>" & _
        "<ReqType>SUB</ReqType>" & _
        "<Symbols>" & _
        """" & UCase$(pContract.ProviderID(mProviderKey)) & """" & _
        "</Symbols>" & _
        "</L1Quotes>" & _
        Chr(255)
Debug.Print "Request L1 quotes: " & qtQuote
mSocket.SendData qtQuote

End Function

Friend Sub unsubscribeFromL1Quotes( _
                ByVal L1QuoteRequestId As Long)
mL1QuotesRequestTable(L1QuoteRequestId).cancelling = True
mL1QuotesRequestTable(L1QuoteRequestId).reuseTime = Now + OneMinute
Set mL1QuotesRequestTable(L1QuoteRequestId).requester = Nothing
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcConnectionString() As String
calcConnectionString = "server=" & mServer & _
                        "; port=" & mPort & _
                        "; providerKey=" & mProviderKey
End Function

Private Sub connectToQT()
Set mSocket = New Winsock
mSocket.Connect mServer, mPort
End Sub

Private Function findHistoricalTickDataRequestEntry(ByVal symbol As String) As HistoricalTickDataRequestEntry
findHistoricalTickDataRequestEntry = mHistoricalTickDataRequestTable(findHistoricalTickDataRequestEntryIndex(symbol))
End Function

Private Function findHistoricalTickDataRequestEntryIndex(ByVal symbol As String) As Long
Dim i As Long
findHistoricalTickDataRequestEntryIndex = -1
For i = 0 To UBound(mHistoricalTickDataRequestTable)
    If mHistoricalTickDataRequestTable(i).symbol = symbol Then
        findHistoricalTickDataRequestEntryIndex = i
        Exit Function
    End If
Next
err.Raise ErrorCodes.IllegalArgumentException, _
            "QTSP.QTAPI::findHistoricalTickDataRequestEntryIndex", _
            "No entry for symbol " & symbol & " exists"
End Function

Private Function findL1QuoteRequestEntryIndex(ByVal symbol As String) As Long
Dim i As Long
findL1QuoteRequestEntryIndex = -1
For i = 0 To UBound(mL1QuotesRequestTable)
    If mL1QuotesRequestTable(i).symbol = symbol Then
        findL1QuoteRequestEntryIndex = i
        Exit Function
    End If
Next
err.Raise ErrorCodes.IllegalArgumentException, _
            "QTSP.QTAPI::findL1QuoteRequestEntryIndex", _
            "No entry for symbol " & symbol & " exists"
End Function

Private Function GetFieldValue( _
                            ByVal fieldID As Long, _
                            ByRef bytesProcessed As Long) As Variant
Dim value As Variant
Dim padding() As Byte

With mFieldTable(fieldID)
    Select Case .DataType
    Case vbString
        If .size <> 0 Then
            mSocket.GetData value, vbString, .size
            bytesProcessed = bytesProcessed + .size
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
        If .size = 8 Then
            mSocket.GetData padding, vbByte + vbArray, 4
            bytesProcessed = bytesProcessed + 4
        End If
    Case vbDouble
        mSocket.GetData value, vbDouble
        bytesProcessed = bytesProcessed + 8
    Case vbByte
        If .size = 1 Then
            mSocket.GetData value, vbByte, 1
            bytesProcessed = bytesProcessed + 1
        Else
            mSocket.GetData value, vbByte + vbArray, .size
            bytesProcessed = bytesProcessed + .size
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
    If mSocket.BytesReceived >= 4 Then
        mSocket.GetData recordID, vbString, 2
        mSocket.GetData recordLength, vbInteger, 2
        bytesProcessed = 4
        If mSocket.BytesReceived < recordLength - bytesProcessed Then
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
    If (mSocket.BytesReceived + bytesProcessed) < recordLength Then
        Exit Function
    End If
    expectingMoreData = False
End If
GotDataForRecord = True
End Function

Private Sub logMessage( _
                ByVal message As String, _
                ByVal logLevel As LogLevels)
If mLogLevel < logLevel Or _
    mLogLevel = LogLevelNone Then Exit Sub
mCommonServiceConsumer.NotifyListeners message, _
                                        TradeBuildSP.StandardListenValueTypes.Log, _
                                        Now
End Sub

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
    logMessage "QT API Error: XML cannot be parsed: " & XMLData, LogLevelLow
    err.Raise LocalErrorCodes.LongJump
End If

Set error = XMLdoc.selectSingleNode("ERROR")
Set errorCode = error.selectSingleNode("ERRORCODE")
Set errorDesc = error.selectSingleNode("ERRORDESC")

logMessage "QT API Error (" & errorCode.Text & "): " & errorDesc.Text, LogLevelLow
err.Raise LocalErrorCodes.LongJump
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
            App.ProductName
    err.Raise LocalErrorCodes.LongJump
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
        .size = CLng("0" & field.getAttribute("Size"))
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

Private Sub ProcessHistoricalDataEnd( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim fieldIdRaw As Integer
Dim fieldID As FieldTypes
Dim symbol As String
Dim index As Long

Do
    mSocket.GetData fieldIdRaw, vbInteger, 2
    fieldID = CLng(fieldIdRaw)
    bytesProcessed = bytesProcessed + 2
    
    Select Case fieldID
    Case FieldTypes.symbol
        symbol = UCase$(GetFieldValue(fieldID, bytesProcessed))
    Case Else
        GetFieldValue fieldID, bytesProcessed
    End Select
Loop While bytesProcessed < recordLength

index = findHistoricalTickDataRequestEntryIndex(symbol)
If Not mHistoricalTickDataRequestTable(index).cancelling Then
    mHistoricalTickDataRequestTable(index).requester.historicalTickDataComplete
End If
releaseHistoricalTickDataRequestEntry symbol
End Sub

Private Sub ProcessHistoricalTick( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)
Dim fieldIdRaw As Integer
Dim fieldID As FieldTypes

Dim timestamp As Date

Dim symbol As String
Dim thisBid As Double
Dim thisAsk As Double
Dim thisLast As Double
Dim thisVolume As Long

Dim bidTick As Tick
Dim askTick As Tick
Dim lastTick As Tick
Dim volumeTick As Tick

Dim index As Long

index = -1 ' to force an error in case no symbol is received

Do
    mSocket.GetData fieldIdRaw, vbInteger, 2
    fieldID = CLng(fieldIdRaw)
    bytesProcessed = bytesProcessed + 2
    
    Select Case fieldID
    Case FieldTypes.symbol
        symbol = UCase$(GetFieldValue(fieldID, bytesProcessed))
        index = findHistoricalTickDataRequestEntryIndex(symbol)
    Case FieldTypes.BidPrice
        thisBid = GetFieldValue(fieldID, bytesProcessed)
    Case FieldTypes.AskPrice
        thisAsk = GetFieldValue(fieldID, bytesProcessed)
    Case FieldTypes.LastPrice
        thisLast = GetFieldValue(fieldID, bytesProcessed)
    Case FieldTypes.volume
        thisVolume = GetFieldValue(fieldID, bytesProcessed)
    Case FieldTypes.timestamp
        Dim timestampString As String
        timestampString = GetFieldValue(fieldID, bytesProcessed)
        timestamp = CDate(mMonths(Mid$(timestampString, 3, 2)) & " " & _
                            Left$(timestampString, 2) & " " & _
                            Mid$(timestampString, 5, 4) & " " & _
                            Mid$(timestampString, 9, 2) & ":" & _
                            Mid$(timestampString, 11, 2) & ":" & _
                            Mid(timestampString, 13, 2))
                            
    Case Else
        GetFieldValue fieldID, bytesProcessed
    End Select

Loop While bytesProcessed < recordLength

If thisBid <> 0 And thisBid <> mHistoricalTickDataRequestTable(index).prevBid Then
    bidTick.TickType = TickTypes.bid
    bidTick.Price = thisBid
    mHistoricalTickDataRequestTable(index).prevBid = thisBid
End If

If thisAsk <> 0 And thisAsk <> mHistoricalTickDataRequestTable(index).prevAsk Then
    askTick.TickType = TickTypes.ask
    askTick.Price = thisAsk
    mHistoricalTickDataRequestTable(index).prevAsk = thisAsk
End If

If thisVolume <> 0 And thisVolume <> mHistoricalTickDataRequestTable(index).prevVolume Then
    If thisVolume > mHistoricalTickDataRequestTable(index).prevVolume Then
        lastTick.TickType = TickTypes.last
        lastTick.Price = thisLast
        lastTick.size = thisVolume - mHistoricalTickDataRequestTable(index).prevVolume
    End If

    volumeTick.TickType = TickTypes.volume
    volumeTick.size = thisVolume
    mHistoricalTickDataRequestTable(index).prevVolume = thisVolume
End If

If mHistoricalTickDataRequestTable(index).cancelling Then Exit Sub

timestamp = timestamp + mHistoricalTickDataRequestTable(index).timestampOffset

If timestamp >= mHistoricalTickDataRequestTable(index).endTimeLocal Then Exit Sub

If timestamp >= mHistoricalTickDataRequestTable(index).startTimeLocal And _
    timestamp < mHistoricalTickDataRequestTable(index).endTimeLocal _
Then
    If bidTick.TickType <> TickTypes.None Then
        bidTick.timestamp = timestamp
        mHistoricalTickDataRequestTable(index).requester.historicalTick bidTick
    ElseIf Not mHistoricalTickDataRequestTable(index).firstInScopeTickProcessed Then
        If mHistoricalTickDataRequestTable(index).prevBid <> 0# Then
            bidTick.TickType = TickTypes.bid
            bidTick.Price = mHistoricalTickDataRequestTable(index).prevBid
            bidTick.timestamp = timestamp
            mHistoricalTickDataRequestTable(index).requester.historicalTick bidTick
        End If
    End If
    If askTick.TickType <> TickTypes.None Then
        askTick.timestamp = timestamp
        mHistoricalTickDataRequestTable(index).requester.historicalTick askTick
    ElseIf Not mHistoricalTickDataRequestTable(index).firstInScopeTickProcessed Then
        mHistoricalTickDataRequestTable(index).firstInScopeTickProcessed = True
        If mHistoricalTickDataRequestTable(index).prevAsk <> 0# Then
            askTick.TickType = TickTypes.ask
            askTick.Price = mHistoricalTickDataRequestTable(index).prevAsk
            askTick.timestamp = timestamp
            mHistoricalTickDataRequestTable(index).requester.historicalTick askTick
        End If
    End If
    If lastTick.TickType <> TickTypes.None Then
        lastTick.timestamp = timestamp
        mHistoricalTickDataRequestTable(index).requester.historicalTick lastTick
    End If
    If volumeTick.TickType <> TickTypes.None Then
        volumeTick.timestamp = timestamp
        mHistoricalTickDataRequestTable(index).requester.historicalTick volumeTick
    End If
    
    mHistoricalTickDataRequestTable(index).firstInScopeTickProcessed = True
End If

End Sub

Private Sub processL1Quote( _
                            ByVal recordLength As Integer, _
                            ByRef bytesProcessed As Long)

Dim fieldIdRaw As Integer
Dim fieldID As FieldTypes

Dim timestamp As Date

Dim symbol As String

Dim bidTick As Tick
Dim askTick As Tick
Dim lastTick As Tick
Dim highTick As Tick
Dim lowTick As Tick
Dim prevCloseTick As Tick
Dim volumeTick As Tick
Dim openInterestTick As Tick

Dim skippedValue As Variant

Dim index As Long

index = -1 ' to force an error in case no symbol is received

' NB: we t assume that the symbol field is always the first

Do
    mSocket.GetData fieldIdRaw, vbInteger, 2
    fieldID = CLng(fieldIdRaw)
    bytesProcessed = bytesProcessed + 2
    
    Select Case fieldID
    Case FieldTypes.symbol
        symbol = UCase$(GetFieldValue(fieldID, bytesProcessed))
        index = findL1QuoteRequestEntryIndex(symbol)
    Case FieldTypes.BidPrice
        mL1QuotesRequestTable(index).bid = GetFieldValue(fieldID, bytesProcessed)
        bidTick.TickType = TickTypes.bid
    Case FieldTypes.bidSize
        mL1QuotesRequestTable(index).bidSize = GetFieldValue(fieldID, bytesProcessed)
        bidTick.TickType = TickTypes.bid
    Case FieldTypes.AskPrice
        mL1QuotesRequestTable(index).ask = GetFieldValue(fieldID, bytesProcessed)
        askTick.TickType = TickTypes.ask
    Case FieldTypes.askSize
        mL1QuotesRequestTable(index).askSize = GetFieldValue(fieldID, bytesProcessed)
        askTick.TickType = TickTypes.ask
    Case FieldTypes.LastPrice
        mL1QuotesRequestTable(index).last = GetFieldValue(fieldID, bytesProcessed)
        lastTick.TickType = TickTypes.last
    Case FieldTypes.lastSize
        mL1QuotesRequestTable(index).lastSize = GetFieldValue(fieldID, bytesProcessed)
        lastTick.TickType = TickTypes.last
    Case FieldTypes.high
        mL1QuotesRequestTable(index).high = GetFieldValue(fieldID, bytesProcessed)
        highTick.TickType = TickTypes.high
    Case FieldTypes.low
        mL1QuotesRequestTable(index).low = GetFieldValue(fieldID, bytesProcessed)
        lowTick.TickType = TickTypes.low
    Case FieldTypes.prevClose
        mL1QuotesRequestTable(index).prevClose = GetFieldValue(fieldID, bytesProcessed)
        prevCloseTick.TickType = TickTypes.prevClose
    Case FieldTypes.volume
        mL1QuotesRequestTable(index).volume = GetFieldValue(fieldID, bytesProcessed)
        volumeTick.TickType = TickTypes.volume
    Case FieldTypes.openInterest
        mL1QuotesRequestTable(index).openInterest = GetFieldValue(fieldID, bytesProcessed)
        openInterestTick.TickType = TickTypes.openInterest
    Case FieldTypes.timestamp
        Dim timestampString As String
        timestampString = GetFieldValue(fieldID, bytesProcessed)
        timestamp = CDate(mMonths(Mid$(timestampString, 3, 2)) & " " & _
                            Left$(timestampString, 2) & " " & _
                            Mid$(timestampString, 5, 4) & " " & _
                            Mid$(timestampString, 9, 2) & ":" & _
                            Mid$(timestampString, 11, 2) & ":" & _
                            Mid(timestampString, 13, 2))
        mL1QuotesRequestTable(index).timestamp = timestamp + mL1QuotesRequestTable(index).timestampOffset
    Case Else
        skippedValue = GetFieldValue(fieldID, bytesProcessed)
        Debug.Print "Skipped field(" & fieldID & "): " & skippedValue
    End Select

Loop While bytesProcessed < recordLength

If mL1QuotesRequestTable(index).cancelling Then Exit Sub

If askTick.TickType <> TickTypes.None Then
    askTick.timestamp = mL1QuotesRequestTable(index).timestamp
    askTick.Price = mL1QuotesRequestTable(index).ask
    askTick.size = mL1QuotesRequestTable(index).askSize
    mL1QuotesRequestTable(index).requester.fireTick askTick
End If
If bidTick.TickType <> TickTypes.None Then
    bidTick.timestamp = mL1QuotesRequestTable(index).timestamp
    bidTick.Price = mL1QuotesRequestTable(index).bid
    bidTick.size = mL1QuotesRequestTable(index).bidSize
    mL1QuotesRequestTable(index).requester.fireTick bidTick
End If
If lastTick.TickType <> TickTypes.None Then
    lastTick.timestamp = mL1QuotesRequestTable(index).timestamp
    lastTick.Price = mL1QuotesRequestTable(index).last
    lastTick.size = mL1QuotesRequestTable(index).lastSize
    mL1QuotesRequestTable(index).requester.fireTick lastTick
End If
If highTick.TickType <> TickTypes.None Then
    highTick.timestamp = mL1QuotesRequestTable(index).timestamp
    highTick.Price = mL1QuotesRequestTable(index).high
    mL1QuotesRequestTable(index).requester.fireTick highTick
End If
If lowTick.TickType <> TickTypes.None Then
    lowTick.timestamp = mL1QuotesRequestTable(index).timestamp
    lowTick.Price = mL1QuotesRequestTable(index).low
    mL1QuotesRequestTable(index).requester.fireTick lowTick
End If
If prevCloseTick.TickType <> TickTypes.None Then
    prevCloseTick.timestamp = mL1QuotesRequestTable(index).timestamp
    prevCloseTick.Price = mL1QuotesRequestTable(index).prevClose
    mL1QuotesRequestTable(index).requester.fireTick prevCloseTick
End If
If volumeTick.TickType <> TickTypes.None Then
    volumeTick.timestamp = mL1QuotesRequestTable(index).timestamp
    volumeTick.size = mL1QuotesRequestTable(index).volume
    mL1QuotesRequestTable(index).requester.fireTick volumeTick
End If
If openInterestTick.TickType <> TickTypes.None Then
    openInterestTick.timestamp = mL1QuotesRequestTable(index).timestamp
    openInterestTick.size = mL1QuotesRequestTable(index).openInterest
    mL1QuotesRequestTable(index).requester.fireTick openInterestTick
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
            App.ProductName
    err.Raise LocalErrorCodes.LongJump
End If

Set login = XMLdoc.selectSingleNode("LOGIN")
Set statusCode = login.selectSingleNode("STATUSCODE")
Select Case statusCode.Text
Case "0"
    ' login successful
    mConnectionState = ConnectionStates.ConnConnected
    RaiseEvent Connected
Case "1"
    ' login failed - bad password
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.PasswordInvalid, _
            "Logon password invalid", _
            App.ProductName
    logMessage "Failed to connect to QuoteTracker: logon password invalid: " & _
                calcConnectionString, _
                LogLevelLow
    RaiseEvent ConnectFailed("logon password invalid", False)
    err.Raise LocalErrorCodes.LongJump
Case Else
    Set statusDesc = XMLdoc.selectSingleNode("STATUSDESC")
    mCommonServiceConsumer.ServiceProviderError SPErrorCodes.UnexpectedError, _
            statusCode.Text & ": " & statusDesc.Text, _
            App.ProductName
    RaiseEvent ConnectFailed("Logon error " & statusCode.Text & ": " & statusDesc.Text, False)
    err.Raise LocalErrorCodes.LongJump
End Select
End Sub

Private Function ProcessRecord() As Boolean
Static recordID As String
Static recordLength As Integer
Static bytesProcessed As Long  ' number of bytes processed in the current record

If Not GotDataForRecord(recordID, recordLength, bytesProcessed) Then
    If recordID <> "" Then Debug.Print "Incomplete " & recordID & " record: missing " & (recordLength - bytesProcessed - mSocket.BytesReceived) & " bytes"
    Exit Function
End If

Select Case recordID
Case "DE"       ' field descriptions
    ProcessFieldDescriptions recordLength, bytesProcessed
    Debug.Print "Processed DE record"
Case "ER"       ' error
    ProcessError recordLength, bytesProcessed
    Debug.Print "Processed ER record"
Case "HT"       ' historical tick
    ProcessHistoricalTick recordLength, bytesProcessed
    Debug.Print "Processed HT record"
Case "HE"       ' end of historical data
    ProcessHistoricalDataEnd recordLength, bytesProcessed
    Debug.Print "Processed HE record"
Case "LO"       ' login response
    ProcessLoginResponse recordLength, bytesProcessed
    Debug.Print "Processed LO record"
Case "L1"       ' level 1 quote update
    processL1Quote recordLength, bytesProcessed
    Debug.Print "Processed S1 record"
Case "S1"       ' level 1 quote
    processL1Quote recordLength, bytesProcessed
    Debug.Print "Processed S1 record"
Case Else
    Dim data As String
    Debug.Print "Processed " & recordID & " record"
    ' get the rest of the message off the socket
    mSocket.GetData data, vbString, recordLength - bytesProcessed
    bytesProcessed = bytesProcessed + Len(data)
End Select

recordID = ""
recordLength = 0
bytesProcessed = 0

mTotalBytesProcessed = mTotalBytesProcessed + bytesProcessed
ProcessRecord = True
End Function

Private Sub ProcessSocketData()

Do While ProcessRecord
        If mSocket.BytesReceived + mTotalBytesProcessed > mTotalBytesInBuffer Then
            mTotalBytesInBuffer = mSocket.BytesReceived + mTotalBytesProcessed
            Debug.Print "Socket data increased: " & mTotalBytesInBuffer
        End If
Loop

End Sub

Private Sub releaseHistoricalTickDataRequestEntry(ByVal symbol As String)
Dim index As Long

index = findHistoricalTickDataRequestEntryIndex(symbol)
mHistoricalTickDataRequestTable(index).endTimeLocal = 0
mHistoricalTickDataRequestTable(index).firstInScopeTickProcessed = False
mHistoricalTickDataRequestTable(index).inUse = False
mHistoricalTickDataRequestTable(index).prevAsk = 0
mHistoricalTickDataRequestTable(index).prevBid = 0
mHistoricalTickDataRequestTable(index).prevVolume = 0
Set mHistoricalTickDataRequestTable(index).requester = Nothing
mHistoricalTickDataRequestTable(index).startTimeLocal = 0
mHistoricalTickDataRequestTable(index).symbol = ""
mHistoricalTickDataRequestTable(index).timestampOffset = 0
End Sub

Private Sub ReleaseSocket()
If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
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

Private Sub RequestLogin()
Dim qtLogin As String
qtLogin = "<LOGIN><PASSWORD>mPassword</PASSWORD>" & _
            "<AppName>TBQTSP</AppName>" & _
            "<AppVer>1</AppVer>" & _
            "</LOGIN>" + _
            Chr(255)

mSocket.SendData qtLogin
End Sub

Private Sub retryConnection()
If mConnectionRetryIntervalSecs <> 0 Then
    If mConnectionTimer Is Nothing Then
        Set mConnectionTimer = New IntervalTimer
        mConnectionTimer.RepeatNotifications = False
        mConnectionTimer.TimerIntervalSecs = mConnectionRetryIntervalSecs
    End If
    mConnectionTimer.StartTimer
End If
End Sub



