VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

' Raised when a successful connection to Tws has been achieved.
Event Connected()

' Raised when connection to Tws has failed and the connection
' retry interval is zero.
Event ConnectFailed(ByVal Description As String, _
                ByVal retrying As Boolean)

' Raised when an attempt to connect to Tws is initiated.
Event Connecting()

' Raised when the connection to Tws is disconnected.
Event ConnectionToTwsClosed( _
                ByVal reconnecting As Boolean)

Event ConnectionToIBClosed()

Event ConnectionToIBRecovered()

' Raised when we Disconnect from Tws
Event Disconnected(ByVal pReason As String, ByVal pForced As Boolean)

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                    As String = "TwsAPI"

Private Const ClientVersion As Long = 47
Private Const ReqdServerVersion As Long = 38

Private Const BaseMarketDataRequestId           As Long = 0
Private Const BaseMarketDepthRequestId          As Long = &H400000
Private Const BaseHistoricalDataRequestId       As Long = &H800000
Private Const BaseExecutionsRequestId           As Long = &H810000
Private Const BaseContractRequestId             As Long = &H1000000
Private Const BaseOrderId                       As Long = &H10000000

Private Const MaxCallersMarketDataRequestId     As Long = BaseMarketDepthRequestId - 1
Private Const MaxCallersMarketDepthRequestId    As Long = BaseHistoricalDataRequestId - BaseMarketDepthRequestId - 1
Private Const MaxCallersHistoricalDataRequestId As Long = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
Private Const MaxCallersExecutionsRequestId     As Long = BaseContractRequestId - BaseExecutionsRequestId - 1
Private Const MaxCallersContractRequestId       As Long = BaseOrderId - BaseContractRequestId - 1
Private Const MaxCallersOrderId                 As Long = &H7FFFFFFF - BaseOrderId - 1

Private Const MIN_SERVER_VER_REAL_TIME_BARS             As Long = 34
Private Const MIN_SERVER_VER_SCALE_ORDERS               As Long = 35
Private Const MIN_SERVER_VER_SNAPSHOT_MKT_DATA          As Long = 35
Private Const MIN_SERVER_VER_SSHORT_COMBO_LEGS          As Long = 35
Private Const MIN_SERVER_VER_WHAT_IF_ORDERS             As Long = 36
Private Const MIN_SERVER_VER_CONTRACT_CONID             As Long = 37
Private Const MIN_SERVER_VER_PTA_ORDERS                 As Long = 39
Private Const MIN_SERVER_VER_FUNDAMENTAL_DATA           As Long = 40
Private Const MIN_SERVER_VER_UNDER_COMP                 As Long = 40
Private Const MIN_SERVER_VER_CONTRACT_DATA_CHAIN        As Long = 40
Private Const MIN_SERVER_VER_SCALE_ORDERS2              As Long = 40
Private Const MIN_SERVER_VER_ALGO_ORDERS                As Long = 41
Private Const MIN_SERVER_VER_EXECUTION_DATA_CHAIN       As Long = 42
Private Const MIN_SERVER_VER_NOT_HELD                   As Long = 44
Private Const MIN_SERVER_VER_SEC_ID_TYPE                As Long = 45
Private Const MIN_SERVER_VER_PLACE_ORDER_CONID          As Long = 46
Private Const MIN_SERVER_VER_REQ_MKT_DATA_CONID         As Long = 47
Private Const MIN_SERVER_VER_REQ_CALC_IMPLIED_VOLAT     As Long = 49
Private Const MIN_SERVER_VER_REQ_CALC_OPTION_PRICE      As Long = 50
Private Const MIN_SERVER_VER_CANCEL_CALC_IMPLIED_VOLAT  As Long = 50
Private Const MIN_SERVER_VER_CANCEL_CALC_OPTION_PRICE   As Long = 50

'================================================================================
' Enums
'================================================================================

Private Enum IdTypes
    IdTypeNone
    IdTypeRealtimeData
    IdTypeMarketDepth
    IdTypeHistoricalData
    IdTypeOrder
    IdTypeContractData
    IdTypeExecution
End Enum

Private Enum InternalErrorCodes
    DataIncomplete = vbObjectError + 4327   ' let's hope nothing else uses this number!
End Enum

Private Enum TwsHistDataDateFormats
    DateFormatString = 1
    DateFormatInteger = 2
End Enum
'================================================================================
' Types
'================================================================================

Private Type StatsEntry
    lastSecondCount         As Long
    lastSecondTime          As Single
    lastPeriodCount         As Long
    lastPeriodTime          As Single
    totalCount              As Long
    totalTime               As Single
    maxSecondCount          As Long
    longestTime             As Single
    shortestTime            As Single
End Type

'================================================================================
' Member variables
'================================================================================

Private mTws                                    As Tws
Private mClientID                               As Long

Private WithEvents mSocket                      As Winsock
Attribute mSocket.VB_VarHelpID = -1
Private mWinsockForm                            As WinsockForm

Private WithEvents mConnectionTimer             As IntervalTimer
Attribute mConnectionTimer.VB_VarHelpID = -1

Private mConnectionRetryIntervalSecs            As Long
Private mDisconnectedByApplication              As Boolean

Private mSocketInputData()                      As Byte

Private mOutputBuffer()                         As Byte
Private mOutputBufferIndex                      As Long

Private mInputBuffer()                          As Byte
Private mInputBufferNextFreeIndex               As Long

Private mInputParseIndex                        As Long
Private mInputParseIndexChk                     As Long
Private mCheckPointNumber                       As Long

Private mServerVersion                          As Long

Private mConnectionState                        As ConnectionStates

Private mNextOrderID                            As Long

Private mTwsLogLevel                            As TwsLogLevels

Private mPerformanceStats(MAX_SOCKET_INMSG)     As StatsEntry

Private WithEvents mPerformanceTimerSecond      As IntervalTimer
Attribute mPerformanceTimerSecond.VB_VarHelpID = -1
Private WithEvents mPerformanceTimerPeriod      As IntervalTimer
Attribute mPerformanceTimerPeriod.VB_VarHelpID = -1

Private mPerformanceElapsedTimer                As ElapsedTimer

Private mPerformanceLogger                      As Logger
Private mSocketLogger                           As Logger

Private mInMessageBuilder                       As StringBuilder

Private mOutMessageBuilder                      As StringBuilder

Private mRetryingConnection                     As Boolean

Private mUsageCount                             As Long

Private mContractDetailsRequestInProgress       As Boolean
Private mGotFirstContractDetails                As Boolean

'================================================================================
' External function declarations
'================================================================================

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)
                            
Private Declare Sub MoveMemory Lib "Kernel32" Alias "RtlMoveMemory" ( _
                            Destination As Any, _
                            source As Any, _
                            ByVal length As Long)
                            
'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

ReDim mInputBuffer(1023) As Byte
ReDim mOutputBuffer(15) As Byte

createLoggers

mConnectionState = ConnectionStates.ConnNotConnected

mNextOrderID = BaseOrderId

mTwsLogLevel = TwsLogLevelError ' this is the Tws default

Set mPerformanceElapsedTimer = New ElapsedTimer

End Sub

Private Sub Class_Terminate()
Debug.Print "TwsAPI terminated"
End Sub

'================================================================================
' mConnectionTimer Event Handlers
'================================================================================

Private Sub mConnectionTimer_TimerExpired()
Const ProcName As String = "mConnectionTimer_TimerExpired"
On Error GoTo Err
Set mConnectionTimer = Nothing
Connect
Exit Sub
Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mPerformanceTimerSecond Event Handlers
'================================================================================

Private Sub mPerformanceTimerPeriod_TimerExpired()
Const ProcName As String = "mPerformanceTimerPeriod_TimerExpired"
Dim i As Long
Dim s As String

On Error GoTo Err

For i = 0 To TwsSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .lastPeriodCount = .lastPeriodCount + .lastSecondCount
        .lastPeriodTime = .lastPeriodTime + .lastSecondTime
        .totalCount = .totalCount + .lastSecondCount
        .totalTime = .totalTime + .lastSecondTime
        
        If .totalCount <> 0 Then
            s = gSocketInMsgTypeToString(i)
            s = s & Format(.lastPeriodCount, " 00000")
            If .lastPeriodCount <> 0 Then
                s = s & Format(.lastPeriodTime / .lastPeriodCount, " 00000.0")
            Else
                s = s & "     0.0"
            End If
            s = s & Format(.totalCount, " 00000000")
            If .totalCount <> 0 Then
                s = s & Format(.totalTime / .totalCount, " 00000.0")
            Else
                s = s & "     0.0"
            End If
            s = s & Format(.maxSecondCount, " 0000")
            s = s & Format(.longestTime, " 00000.0")
            s = s & Format(.shortestTime, " 00000.0")
            mPerformanceLogger.Log LogLevelMediumDetail, s
        End If
        
        .lastPeriodCount = 0
        .lastPeriodTime = 0
        .lastSecondCount = 0
        .lastSecondTime = 0
        .maxSecondCount = 0
    End With
Next
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mPerformanceTimerSecond_TimerExpired()
Const ProcName As String = "mPerformanceTimerSecond_TimerExpired"
Dim i As Long

On Error GoTo Err

For i = 0 To TwsSocketInMsgTypes.MAX_SOCKET_INMSG
    With mPerformanceStats(i)
        .lastPeriodCount = .lastPeriodCount + .lastSecondCount
        .lastPeriodTime = .lastPeriodTime + .lastSecondTime
        If .lastSecondTime > .maxSecondCount Then .maxSecondCount = .lastSecondCount
        .totalCount = .totalCount + .lastSecondCount
        .totalTime = .totalTime + .lastSecondTime
        .lastSecondCount = 0
        .lastSecondTime = 0
    End With
Next
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Socket Event Handlers
'================================================================================

Private Sub mSocket_Close()
Const ProcName As String = "mSocket_Close"
On Error GoTo Err
mSocket.Close
handleTwsDisconnection
Exit Sub
Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSocket_Connect()
Const ProcName As String = "mSocket_Connect"
On Error GoTo Err

If mRetryingConnection Then
    ' this means we have been waiting for Tws to start. If we try to start tickers
    ' too soon after Tws accepts socket connections, Tws gives a 'max number of
    ' tickers reached' error. So we'll cancel this connection and connect again in
    ' 10 seconds
    logMessage "Tws is now available - connecting again in 10 seconds to allow Tws Time to initialise", "mSocket_Connect"
    mSocket.Close
    Set mConnectionTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = False
    Exit Sub
End If

mInputBufferNextFreeIndex = 0
mOutputBufferIndex = 0
mInputParseIndex = 0
mCheckPointNumber = 0

logMessage "Connecting to Tws: negotiating versions and client id", pProcName:="mSocket_Connect"
addString ClientVersion, "Client version"
send
addString mClientID, "Client id"
send
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSocket_DataArrival( _
                ByVal bytesTotal As Long)
Const ProcName As String = "mSocket_DataArrival"
Dim bytesReceived As Long
Static processingMessages As Boolean

On Error GoTo Err

If mSocket.state <> StateConstants.sckConnected Then Exit Sub

mSocket.GetData mSocketInputData, vbArray + vbByte
bytesReceived = UBound(mSocketInputData) + 1
'Debug.Print "Received " & bytesReceived & " from socket: " & showBuffer(mSocketInputData, bytesReceived)

If mSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    mSocketLogger.Log LogLevels.LogLevelMediumDetail, _
                        showBuffer("In buf: ", mSocketInputData)
End If

If mInputBufferNextFreeIndex = mInputParseIndex Then
    ' no unprocessed data so start at the beginning of the buffer again
    'Debug.Print "Reverting to start of buffer (no unprocessed data)"
    mInputBufferNextFreeIndex = 0
    mInputParseIndex = 0
    mInputParseIndexChk = 0
End If

If mInputBufferNextFreeIndex + bytesReceived > UBound(mInputBuffer) + 1 Then
    If bytesReceived <= (UBound(mInputBuffer) + 1 - mInputBufferNextFreeIndex + mInputParseIndex) Then
        ' copy the existing contents down to the start of the buffer
        'Debug.Print "Moving unprocessed data (" & mInputBufferNextFreeIndex - mInputParseIndex & " bytes) from " & mInputParseIndex & " to start of buffer"
        MoveMemory mInputBuffer(0), mInputBuffer(mInputParseIndex), mInputBufferNextFreeIndex - mInputParseIndex
        mInputBufferNextFreeIndex = mInputBufferNextFreeIndex - mInputParseIndex
        mInputParseIndex = 0
        mInputParseIndexChk = 0
    Else
        If mSocketLogger.IsLoggable(LogLevelHighDetail) Then
            mSocketLogger.Log LogLevels.LogLevelHighDetail, _
                            "Extending socket buffer by " & IIf(bytesReceived > 1024, bytesReceived, 1024) & " bytes"
        End If
        ReDim Preserve mInputBuffer(UBound(mInputBuffer) + IIf(bytesReceived > 1024, bytesReceived, 1024)) As Byte
    End If
End If
'Debug.Print "Appending socket message to buffer at " & mInputBufferNextFreeIndex
CopyMemory mInputBuffer(mInputBufferNextFreeIndex), mSocketInputData(0), bytesReceived
mInputBufferNextFreeIndex = mInputBufferNextFreeIndex + bytesReceived

If mInputBufferNextFreeIndex <> mInputParseIndex Then
    If processingMessages Then
        ' this is a re-entrant call to DataArrival. We don't want to call checkMessages
        ' here, because the current call to it will process the data we've just
        ' added to the buffer
        mSocketLogger.Log LogLevels.LogLevelNormal, _
                        "Re-entrant call to mSocket_DataArrival"
    Else
        processingMessages = True
        checkmessages
        processingMessages = False
    End If
End If

Exit Sub
Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSocket_Error(ByVal number As Integer, _
                        Description As String, _
                        ByVal Scode As Long, _
                        ByVal source As String, _
                        ByVal HelpFile As String, _
                        ByVal HelpContext As Long, _
                        CancelDisplay As Boolean)
Const ProcName As String = "mSocket_Error"
Dim errorNum As MSWinsockLib.ErrorConstants

On Error GoTo Err
errorNum = number
Select Case errorNum
'Case sckInvalidPropertyValue

'Case sckGetNotSupported

'Case sckSetNotSupported

'Case sckOutOfMemory

'Case sckBadState

'Case sckInvalidArg

'Case sckSuccess

'Case sckUnsupported

'Case sckInvalidOp

'Case sckOutOfRange

'Case sckWrongProtocol

'Case sckOpCanceled

'Case sckInvalidArgument

'Case sckWouldBlock

'Case sckInProgress

'Case sckAlreadyComplete

'Case sckNotSocket

'Case sckMsgTooBig

'Case sckPortNotSupported

'Case sckAddressInUse

Case sckAddressNotAvailable, _
    sckNetworkSubsystemFailed, _
    sckNetworkUnreachable, _
    sckNetReset, _
    sckConnectAborted, _
    sckConnectionRefused, _
    sckHostNotFound, _
    sckHostNotFoundTryAgain, _
    sckTimedout, _
    sckConnectionReset
    
    ReleaseSocket
    If mConnectionState = ConnectionStates.ConnNotConnected Then
        If mConnectionRetryIntervalSecs = 0 Then
            logMessage "Failed to connect to Tws", _
                        pMsgQualifier:=Description & ": " & _
                                        ConnectionString, _
                        pProcName:=ProcName
            
            RaiseEvent ConnectFailed(Description, False)
        Else
            logMessage "Failed to connect to Tws - retrying in " & mConnectionRetryIntervalSecs & " seconds", _
                        pMsgQualifier:=Description & ": " & _
                                        ConnectionString, _
                        pProcName:=ProcName
            
            RaiseEvent ConnectFailed(Description, True)
            retryConnection
        End If
    Else
        logMessage "Socket error " & number & ": " & _
                    Description & ": " & _
                    ConnectionString, _
                    ProcName
                    
        handleTwsDisconnection
                    
    End If

'Case sckNoBufferSpace

'Case sckAlreadyConnected

'Case sckNotConnected

'Case sckSocketShutdown

'Case sckNotInitialized

'Case sckNonRecoverableError

'Case sckNoData
Case Else
    logMessage "Socket error " & number & ": " & _
                Description & ": " & _
                ConnectionString, _
                ProcName, _
                pLogLevel:=LogLevelSevere
    RaiseEvent ConnectFailed(Description, False)
End Select

Exit Sub
Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Get ClientID() As Long
ClientID = mClientID
End Property

Friend Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
Const ProcName As String = "ConnectionRetryIntervalSecs"
On Error GoTo Err

If Value < 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value cannot be negative"
If mConnectionRetryIntervalSecs > 0 And Value < mConnectionRetryIntervalSecs Then mConnectionRetryIntervalSecs = Value
mConnectionRetryIntervalSecs = Value
If mConnectionRetryIntervalSecs = 0 Then Set mConnectionTimer = Nothing

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Get ConnectionState() As ConnectionStates
ConnectionState = mConnectionState
End Property

Friend Property Get ConnectionString() As String
ConnectionString = "server=" & mTws.Server & _
                        " port=" & mTws.Port & _
                        " client Id=" & mClientID
End Property

Friend Property Get MaxContractRequestId() As Long
MaxContractRequestId = BaseOrderId - BaseContractRequestId - 1
End Property

Friend Property Get MaxExecutionsRequestId() As Long
MaxExecutionsRequestId = BaseContractRequestId - BaseExecutionsRequestId - 1
End Property

Friend Property Get MaxHistoricalDataRequestId() As Long
MaxHistoricalDataRequestId = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
End Property

Friend Property Get MaxMarketDataRequestId() As Long
MaxMarketDataRequestId = BaseMarketDepthRequestId - 1
End Property

Friend Property Get MaxMarketDepthRequestId() As Long
MaxMarketDepthRequestId = BaseHistoricalDataRequestId - BaseMarketDepthRequestId - 1
End Property

Friend Property Get MaxOrderId() As Long
MaxOrderId = &H7FFFFFFF
End Property

Friend Property Get NextOrderId() As Long
NextOrderId = mNextOrderID
mNextOrderID = mNextOrderID + 1
End Property

Public Property Get SerialiseContractDetailsRequests() As Boolean
SerialiseContractDetailsRequests = (mServerVersion < MIN_SERVER_VER_CONTRACT_DATA_CHAIN)
End Property

Friend Property Get serverVersion() As Long
serverVersion = mServerVersion
End Property

Friend Property Get Tws() As Tws
Set Tws = mTws
End Property

Friend Property Let TwsLogLevel(ByVal Value As TwsLogLevels)
mTwsLogLevel = Value
End Property

Friend Property Get TwsLogLevel() As TwsLogLevels
TwsLogLevel = mTwsLogLevel
End Property

Friend Property Get UsageCount() As Long
UsageCount = mUsageCount
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub CancelHistoricalData( _
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalData"
On Error GoTo Err

cancelHistDataRequest pRequestId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CancelMarketData(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

cancelMarketDataEx pTickerId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CancelMarketDepth(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const Version = 1

addMsgId CANCEL_MKT_DEPTH
addString Version, "Version"
addString getTwsMarketDepthRequestIdFromCallersRequestId(pTickerId), "Ticker id"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub cancelNewsBulletins()
Const ProcName As String = "cancelNewsBulletins"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const Version = 1

addMsgId CANCEL_NEWS_BULLETINS
addString Version, "Version"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub CancelOrder(ByVal pOrderId As Long)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const Version = 1
addMsgId CANCEL_ORDER
addString Version, "Version"
addString pOrderId, "Order id"
send True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Connect()
Const ProcName As String = "Connect"
On Error GoTo Err

mDisconnectedByApplication = False

If mConnectionState <> ConnectionStates.ConnNotConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Already connected"
End If

createsocket

logMessage "Connecting to Tws: " & ConnectionString, "Connect"
mSocket.Connect mTws.Server, mTws.Port

mConnectionState = ConnectionStates.ConnConnecting
RaiseEvent Connecting

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub DecrementUsageCount()
mUsageCount = mUsageCount - 1
End Sub

Friend Sub Disconnect( _
                ByVal reason As String, _
                ByVal forced As Boolean)
Const ProcName As String = "Disconnect"
On Error GoTo Err

If mConnectionState = ConnectionStates.ConnNotConnected Then Exit Sub

mDisconnectedByApplication = True
If Not mConnectionTimer Is Nothing Then mConnectionTimer.StopTimer
mConnectionState = ConnectionStates.ConnNotConnected

ReleaseSocket
logMessage "Disconnected from: " & ConnectionString & ": " & reason, ProcName
handleTwsDisconnection
If Not forced Then RaiseEvent Disconnected(reason, forced)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub IncrementUsageCount()
mUsageCount = mUsageCount + 1
End Sub

Friend Sub Initialise( _
                ByVal pTws As Tws, _
                ByVal pClientId As Long)
Set mTws = pTws
mClientID = pClientId
End Sub

Friend Sub PlaceOrder( _
                ByVal pOrder As TwsOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim i As Long

If mConnectionState <> ConnectionStates.ConnConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"

Const Version = 26
addMsgId PLACE_ORDER
addString Version, "Version"

If pOrder.OrderId = 0 Then pOrder.OrderId = NextOrderId
If pOrder.OrderId < BaseOrderId Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Order id must not be less than " & BaseOrderId
If pOrder.OrderId > mNextOrderID Then mNextOrderID = pOrder.OrderId + 1

addString pOrder.OrderId, "Order id"

' send contract fields
With pContract
    If mServerVersion >= MIN_SERVER_VER_PLACE_ORDER_CONID Then addString .ConId, "Con id"
    addString .Symbol, "Symbol"
    addString gTwsSecTypeToShortString(.Sectype), "Sectype"
    addString .Expiry, "Expiry"
    addDouble .Strike, "Strike"
    addString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then addString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier" ' multiplier - this is the number of units of the underlying
                                                            ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    addString .Exchange, "Exchange"
    If mServerVersion >= 14 Then addString .PrimaryExch, "Primary Exchange"    ' only relevant when
                                                        ' Exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in Tws)
    addString .CurrencyCode, "Currency"
    If mServerVersion >= 2 Then addString .LocalSymbol, "Local Symbol"
End With

With pOrder
    
    ' send main Order fields
    addString gTwsOrderActionToString(.Action), "Action"
    addString .TotalQuantity, "Quantity"
    addString gTwsOrderTypeToString(.OrderType), "Order type"
    If mServerVersion <= 10 And _
        (UCase$(pContract.Exchange) = "LIFFE" Or _
        UCase$(pContract.Exchange) = "LSE") _
    Then
        addDouble .LmtPrice / 100, "Price"
        addDouble .AuxPrice / 100, "Aux price"
    Else
        addDouble .LmtPrice, "Price"
        addDouble .AuxPrice, "Aux price"
    End If
    
    ' send extended Order fields
    addString gTwsOrderTIFToString(.Tif), "TIF"
    addString .OcaGroup, "Oca Group"
    addString .Account, "Account"
    addString .OpenClose, "OpenClose"
    addString .Origin, "Origin"
    addString .OrderRef, "Order ref"
    addBoolean .Transmit, "Transmit"
    If mServerVersion >= 4 Then addString .ParentId, "Parent id"
    If mServerVersion >= 5 Then
        addBoolean .BlockOrder, "Block Order"
        addBoolean .SweepToFill, "Sweep to fill"
        addString .DisplaySize, "Display Size"
        addString stopTriggerMethodToString(.TriggerMethod), "Trigger method"
        If mServerVersion < 38 Then
            addBoolean False, "Outside RTH"
        Else
            addBoolean .OutsideRth, "Outside RTH"
        End If
    End If
    If mServerVersion >= 7 Then addBoolean .Hidden, "Hidden"
End With

' Send combo legs for BAG requests
With pContract
    If mServerVersion >= 8 And .Sectype = TwsSecTypeCombo Then
        If .ComboLegs.Count = 0 Then
            addString 0, "Combo legs count"
        Else
            addString .ComboLegs.Count, "Combo legs count"
            Dim comboLeg As TwsComboLeg
            For Each comboLeg In .ComboLegs
                With comboLeg
                    addString .ConId, "Leg " & i & " Con id"
                    addString .Ratio, "Leg " & i & " Ration"
                    addString gTwsOrderActionToString(.Action), "Leg " & i & " Action"
                    addString .Exchange, "Leg " & i & " Exchange"
                    addString legOpenCloseToString(.OpenClose), "Leg " & i & " Open/close"
                End With
            Next
        End If
    End If
End With

With pOrder
    If mServerVersion >= 9 Then addString "", "Shares Allocation"
    
    If mServerVersion >= 10 Then addString .DiscretionaryAmt, "Discretionary amount"
    
    If mServerVersion >= 11 Then addString .GoodAfterTime, "Good after Time"
    
    If mServerVersion >= 12 Then addString .GoodTillDate, "Good till date"
    
    If mServerVersion >= 13 Then
        addString .FaGroup, "FAGroup"
        addString .FaMethod, "FAMethod"
        addString .FaPercentage, "FAPercentage"
        addString .FaProfile, "FAProfile"
    End If
    
    If mServerVersion >= 18 Then 'institutional short sale slot fields.
        addString .ShortSaleSlot, " Short Sale Slot"           ' 0 only for retail, 1 or 2 only for institution.
        addString .DesignatedLocation, "Designated Location"       ' only populate when shortSaleSlot = 2.
    End If
    
    If mServerVersion >= 19 Then
        addString .OcaType, "Oca type"
        If mServerVersion < 38 Then
            addBoolean False, "RTH only" ' (now combined with ignoreRTH)
        End If
        addBoolean .Rule80A, "Rule 80A"
        addString .SettlingFirm, "Settling firm"
        addBoolean .AllOrNone, "All or none"
        addString .MinQty, "Minimum quantity"
        addString .PercentOffset, "Percent Offset"
        addBoolean .ETradeOnly, "E-trade only"
        addBoolean .FirmQuoteOnly, "Firm quote only"
        addString .NbboPriceCap, "NBBO price cap"
        addString .AuctionStrategy, "Auction strategy"
        addDouble .StartingPrice, "Starting price"
        addDouble .StockRefPrice, "Stock ref price"
        addDouble .Delta, "Delta"
        
        If mServerVersion = 26 And .OrderType = TwsOrderTypeVol Then
            ' Volatility orders had specific watermark price attribs in server version 26
            addDouble MaxDouble, "Stock range lower"
            addDouble MaxDouble, "Stock range upper"
        Else
            addDouble .StockRangeLower, "Stock range lower"
            addDouble .StockRangeUpper, "Stock range upper"
        End If
    End If
    
    If mServerVersion >= 22 Then addBoolean .OverridePercentageConstraints, "Override percentage constraints"

    If mServerVersion >= 26 Then ' Volatility orders
        addDouble .Volatility, "Volatility"
        addString .VolatilityType, "Volatility type"
        If mServerVersion < 28 Then
            addBoolean (.DeltaNeutralOrderType = TwsOrderTypeMarket), "Delta neutral Order type"
        Else
            addString gTwsOrderTypeToString(.DeltaNeutralOrderType), "Delta neutral Order type)"
            addDoubleMax .DeltaNeutralAuxPrice, "Delta neutral aux price"
        End If
        
        addBoolean .ContinuousUpdate, "Continuous update"
        
        If mServerVersion = 26 And .OrderType = TwsOrderTypeVol Then
            ' Volatility orders had specific watermark price attribs in server version 26
            addDouble .StockRangeLower, "Stock range lower"
            addDouble .StockRangeUpper, "Stock range upper"
        Else
            addDouble MaxDouble, "Stock range lower"
            addDouble MaxDouble, "Stock range upper"
        End If
        
        addLongMax .ReferencePriceType, "Reference price type"
    End If
    
    If mServerVersion >= 30 Then addDoubleMax .TrailStopPrice, "Trail stop price"

    If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS) Then
        If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS2) Then
            addLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
            addLongMax .ScaleSubsLevelSize, "ScaleSubsLevelSize"
        
        Else
            addString "", "ScaleSubsLevelSize"
            addLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
            
        End If
        addDoubleMax .ScalePriceIncrement, "ScalePriceIncrement"
    End If
    
    If (mServerVersion >= MIN_SERVER_VER_PTA_ORDERS) Then
        addString .ClearingAccount, "Clearing account"
        addString .ClearingIntent, "Clearing intent"
    End If

    If (mServerVersion >= MIN_SERVER_VER_NOT_HELD) Then
        addBoolean .NotHeld, "Not held"
    End If

    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
        If pContract.UnderComp.ConId <> 0 Then
            addBoolean True, "Under comp"
            addString pContract.UnderComp.ConId, "Under comp conid"
            addDouble pContract.UnderComp.Delta, "Under comp delta"
            addDouble pContract.UnderComp.Price, "Under comp price"
        Else
            addBoolean False, "Under comp"
        End If
    End If

    If (mServerVersion >= MIN_SERVER_VER_ALGO_ORDERS) Then
        addString .AlgoStrategy, "Algo strategy"
        If .AlgoStrategy <> "" Then
            Dim algoParamsCount  As Long
            
            On Error Resume Next
            algoParamsCount = UBound(.AlgoParams) + 1
            On Error GoTo Err
             
            addString algoParamsCount, "Algo params count"
            If algoParamsCount > 0 Then
                Dim lAlgoParams() As TwsTagValue
                lAlgoParams = pOrder.AlgoParams
                For i = 0 To algoParamsCount - 1
                    addString lAlgoParams(i).Tag, "Tag" & i
                    addString lAlgoParams(i).Value, "value" & i
                Next
            End If
        End If
    End If

    If (mServerVersion >= MIN_SERVER_VER_WHAT_IF_ORDERS) Then
        addBoolean .WhatIf, "WhatIf"
    End If
End With

send True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
                    
Friend Sub ReplaceFA(ByVal DataType As FADataTypes, ByVal xml As String)
Const ProcName As String = "ReplaceFA"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already connected"
End If

Const Version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Server version does not support ReplaceFA"
End If

addMsgId REPLACE_FA
addString Version, "Version"
addString DataType, "Data type"
addString xml, "XML"

send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)
Const ProcName As String = "RequestAccountUpdates"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already connected"
End If

Const Version = 2

addMsgId REQ_ACCT_DATA
addString Version, "Version"
addBoolean subscribe, "Subscribe"

If mServerVersion >= 9 Then addString acctCode, "Account code"

send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestAllOpenOrders()
Const ProcName As String = "RequestAllOpenOrders"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Not connected"
End If
                            

Const Version = 1

addMsgId REQ_ALL_OPEN_ORDERS
addString Version, "Version"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestAutoOpenOrders(ByVal autoBind As Boolean)
Const ProcName As String = "RequestAutoOpenOrders"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const Version = 1

addMsgId REQ_AUTO_OPEN_ORDERS
addString Version, "Version"
addBoolean autoBind, "Autobind"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'Friend Sub RequestContract( _
'                ByVal ContractSpec As ContractSpecifier, _
'                ByVal Requester As zzContractDetailsRequester)
'Dim requestQueueEntry As TWSContractDetailsRequestQueueEntry
'
'Const ProcName As String = "RequestContract"
'On Error GoTo Err
'
'Set requestQueueEntry.ContractSpec = ContractSpec
'Set requestQueueEntry.Requester = Requester
'
'Select Case ContractSpec.Sectype
'Case SecurityTypes.SecTypeCombo
'Case SecurityTypes.SecTypeCash
'Case SecurityTypes.SecTypeFuture
'Case SecurityTypes.SecTypeFuturesOption
'Case SecurityTypes.SecTypeIndex
'Case SecurityTypes.SecTypeOption
'Case SecurityTypes.SecTypeStock
'Case Else
'    Requester.ContractSpecifierInvalid requestQueueEntry, "security type must be supplied"
'    Exit Sub
'End Select
'
'If ContractSpec.LocalSymbol = "" And _
'    ContractSpec.Symbol = "" _
'Then
'    Requester.ContractSpecifierInvalid requestQueueEntry, "either Symbol or local Symbol must be supplied"
'    Exit Sub
'End If
'
'mTwsContractDetailsRequestQueue.AddToEnd requestQueueEntry
'
'If mConnectionState = ConnConnected Then submitNextContractDetailsRequest
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Sub

Friend Sub RequestContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

reqContractDetails pRequestId, pContract

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestExecutions( _
                ByVal pRequestId As Long, _
                ByVal filter As TwsExecutionFilter)
Dim theFilter As TwsExecutionFilter

Const ProcName As String = "RequestExecutions"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const Version = 3

addMsgId REQ_EXECUTIONS
addString Version, "Version"

If mServerVersion >= MIN_SERVER_VER_EXECUTION_DATA_CHAIN Then addString getTwsExecutionsRequestIdFromCallersRequestId(pRequestId), "ReqId"

If mServerVersion >= 9 Then
    If filter Is Nothing Then
        Set theFilter = New TwsExecutionFilter
    Else
        Set theFilter = filter
    End If
    
    With theFilter
        addString .ClientID, "Client id"
        addString .AccountCode, "Account code"
        If .Time = Int(.Time) Then
            addString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd"), ""), "Fill Time"
        Else
            addString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd-hh\:nn\:ss"), ""), "Fill Time"
        End If
        addString .Symbol, "Symbol"
        addString gTwsSecTypeToShortString(.Sectype), "Sec type"
        addString .Exchange, "Exchange"
        addString gTwsOrderActionToString(.Action), "Action"
    End With
End If

send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestFA(ByVal DataType As FADataTypes)
Const ProcName As String = "RequestFA"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const Version = 1

If mServerVersion < 13 Then
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Server version does not support RequestFA"
End If

addMsgId REQ_FA
addString Version, "Version"
addString DataType, "Data type"

send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestManagedAccounts()
Const ProcName As String = "RequestManagedAccounts"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
                            ProjectName & "." & ModuleName & ":" & ProcName, _
                "Not connected"
End If

Const Version = 1

addMsgId REQ_MANAGED_ACCTS
addString Version, "Version"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"

reqMarketData pTickerId, pContract, pGenericTicks, pSnapshot

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"

reqMarketDepth pTickerId, pContract, pNumberOfRows

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As HistoricalDataRequest)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

reqHistoricalData pRequestId, pRequest

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestNewsBulletins(allMsgs As Boolean)
Const ProcName As String = "RequestNewsBulletins"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"
End If

Const Version = 1

addMsgId REQ_NEWS_BULLETINS
addString Version, "Version"
addString allMsgs, "All messages"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RequestOpenOrders()
Const ProcName As String = "RequestOpenOrders"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"
End If

Const Version = 1

addMsgId REQ_OPEN_ORDERS
addString Version, "Version"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addBoolean( _
                ByVal val As Boolean, _
                ByRef fieldName As String)
Const ProcName As String = "addBoolean"
On Error GoTo Err

addString IIf(val, 1, 0), fieldName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addByte(ByVal Value As Integer)
Const ProcName As String = "addByte"
On Error GoTo Err

If mOutputBufferIndex > UBound(mOutputBuffer) Then
    ReDim Preserve mOutputBuffer(2 * (UBound(mOutputBuffer) + 1) - 1) As Byte
End If
mOutputBuffer(mOutputBufferIndex) = Value
mOutputBufferIndex = mOutputBufferIndex + 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addDouble( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "addDouble"
On Error GoTo Err

addString DoubleToString(data), fieldName

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addDoubleMax( _
                ByVal data As Double, _
                ByRef fieldName As String)
Const ProcName As String = "addDoubleMax"
On Error GoTo Err

If data = MaxDouble Then
    addString "", fieldName
Else
    addString DoubleToString(data), fieldName
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addLongMax( _
                ByVal data As Long, _
                ByRef fieldName As String)
Const ProcName As String = "addDoubleMax"
On Error GoTo Err

If data = MaxLong Then
    addString "", fieldName
Else
    addString DoubleToString(data), fieldName
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addString( _
                ByVal data As String, _
                ByRef fieldName As String)
Dim i As Long

Const ProcName As String = "addString"
On Error GoTo Err

If mOutMessageBuilder Is Nothing Then
    Set mOutMessageBuilder = CreateStringBuilder
    mOutMessageBuilder.Append "OUT: "
End If

mOutMessageBuilder.Append fieldName
mOutMessageBuilder.Append "="
mOutMessageBuilder.Append data
mOutMessageBuilder.Append ";"

For i = 1 To Len(data)
    addByte Asc(Mid$(data, i, 1))
Next
addByte 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function adjustPrice( _
                ByVal Price As Double, _
                ByVal pCurrencyCode As String) As Double
Const ProcName As String = "adjustPrice"
On Error GoTo Err

If pCurrencyCode = "GBP" And _
    mServerVersion < 23 _
Then
    adjustPrice = Price * 100
Else
    adjustPrice = Price
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub addMsgId( _
                ByVal id As TwsSocketOutMsgTypes)
Const ProcName As String = "addMsgId"
On Error GoTo Err

addString id, "Msg Id (" & gOutputMessageIdToString(id) & ")"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub cancelHistDataRequest( _
                ByVal pRequestId As Long)
Const ProcName As String = "cancelHistDataRequest"
On Error GoTo Err

If mServerVersion < 24 Then
    ' cancel not supported by Tws
    Exit Sub
End If

If mConnectionState = ConnConnected Then
    Const Version As Long = 1

    addMsgId CANCEL_HISTORICAL_DATA
    addString Version, "Version"
    addString getTwsHistRequestIdFromCallersRequestId(pRequestId), "Request id"
    send
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub cancelMarketDataEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mConnectionState <> ConnectionStates.ConnConnected Then Exit Sub

Const Version = 2

addMsgId CANCEL_MKT_DATA
addString Version, "Version"
addString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub checkmessages()
Const ProcName As String = "checkmessages"
On Error GoTo Err

Static msgId As TwsSocketInMsgTypes
Static Version As Long
Static startIndex As Long

Static prevMsgID As TwsSocketInMsgTypes
Static prevVersion As Long
Static prevStartIndex As Long

Static retryCount As Long
Static eventCount As Long
Static mElapsedTimer As ElapsedTimer
Dim s As String

If mElapsedTimer Is Nothing Then
    Set mElapsedTimer = New ElapsedTimer
    mElapsedTimer.StartTiming
End If

If mConnectionState <> ConnectionStates.ConnConnected Then
    ' first data received is plain Server version number
    If mCheckPointNumber = 0 Then
        createInMessageBuilder
        mServerVersion = getLong("Server version")
        checkpointBuffer
        mCheckPointNumber = 100
        'Debug.Print "Set checkpoint number=" & 100
    End If
    
    ' in some circumstances, some versions of Tws send malformed
    ' ERR_MSG messages before sending the version number, so we
    ' need to skip these if they occur. (They are malformed in that
    ' they don't include the id or error code.)
    
    Do While mServerVersion = TwsSocketInMsgTypes.ERR_MSG
        If mCheckPointNumber = 100 Then
            getLong ("Invalid version")
            checkpointBuffer
            mCheckPointNumber = 101
        End If
        If mCheckPointNumber = 101 Then
            getString "Invalid errormsg"
            checkpointBuffer
            mCheckPointNumber = 102
        End If
        
        logSocketInputMessage
        createInMessageBuilder
        
        If mCheckPointNumber = 102 Then
            mServerVersion = getLong("Server version")
            checkpointBuffer
            mCheckPointNumber = 100
        End If
    Loop
    
    logSocketInputMessage
    createInMessageBuilder
    
    If mCheckPointNumber = 100 Then
        If mServerVersion >= 20 Then
            logMessage "Tws Time at connection", pMsgQualifier:=getString("Tws Time"), pProcName:="checkmessages"
        End If
        checkpointBuffer
        mCheckPointNumber = 200
        'Debug.Print "Set checkpoint number=" & 200
    End If
    
    If mServerVersion < ReqdServerVersion Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Tws is out of date and needs to be upgraded to at least version 884"
    End If
    
    logSocketInputMessage
    
    mConnectionState = ConnectionStates.ConnReSynching
    logMessage "Connection to Tws successful", pProcName:="checkmessages"
    
    ' set the log level in Tws
    setTwsLogLevel mTwsLogLevel
    
    mConnectionState = ConnectionStates.ConnConnected
    RaiseEvent Connected
    
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000
End If

Do While mInputParseIndex < mInputBufferNextFreeIndex
    
    If mCheckPointNumber = 1000 Then
        Dim tmpStartIndex As Long
        tmpStartIndex = startIndex
        startIndex = mInputParseIndex
        prevMsgID = msgId
        
        createInMessageBuilder
    
        msgId = getLong("Msg id")
        mInMessageBuilder.Append "("
        mInMessageBuilder.Append gInputMessageIdToString(msgId)
        mInMessageBuilder.Append ") "
        
        prevStartIndex = tmpStartIndex
        checkpointBuffer
        mCheckPointNumber = 1100
        'Debug.Print "Set checkpoint number=" & 1100 & ": msgID=" & msgID
    End If
    If mCheckPointNumber = 1100 Then
        prevVersion = Version
        Version = getLong("Version")
        checkpointBuffer
        mCheckPointNumber = 1200
        'Debug.Print "Set checkpoint number=" & 1200 & ": version=" & version
    End If
    
    mPerformanceElapsedTimer.StartTiming
    
    Select Case msgId
    Case TwsSocketInMsgTypes.TICK_PRICE
        'Debug.Print "getTickPrice"
        getTickPrice Version
    Case TwsSocketInMsgTypes.TICK_SIZE
        'Debug.Print "getTickSize "
        getTickSize Version
    Case TwsSocketInMsgTypes.ORDER_STATUS
        'Debug.Print "getOrderStatus"
        getOrderStatus Version
    Case TwsSocketInMsgTypes.ERR_MSG
        'Debug.Print "getErrorMsg "
        getErrorMsg Version
    Case TwsSocketInMsgTypes.OPEN_ORDER
        'Debug.Print "getOpenOrder "
        getOpenOrder Version
    Case TwsSocketInMsgTypes.ACCT_Value
        'Debug.Print "getAccountValue "
        getAccountValue Version
    Case TwsSocketInMsgTypes.PORTFOLIO_Value
        'Debug.Print "getPortfolioValue "
        getPortfolioValue Version
    Case TwsSocketInMsgTypes.ACCT_UPDATE_TIME
        'Debug.Print "getAccountTime "
        getAccountTime Version
    Case TwsSocketInMsgTypes.NEXT_VALID_ID
        'Debug.Print "getNextValidId"
        getNextValidId Version
    Case TwsSocketInMsgTypes.CONTRACT_DATA
        'Debug.Print "getContractData"
        getContractData Version
    Case TwsSocketInMsgTypes.EXECUTION_DATA
        'Debug.Print "getExecutionData "
        getExecutionData Version
    Case TwsSocketInMsgTypes.MARKET_DEPTH
        'Debug.Print "getMarketDepth"
        getMarketDepth Version
    Case TwsSocketInMsgTypes.MARKET_DEPTH_L2
        'Debug.Print "getMarketDepthL2 "
        getMarketDepthL2 Version
    Case TwsSocketInMsgTypes.NEWS_BULLETINS
        'Debug.Print "getNewsBulletins "
        getNewsBulletins Version
    Case TwsSocketInMsgTypes.MANAGED_ACCTS
        'Debug.Print "getManagedAccounts "
        getManagedAccounts Version
    Case TwsSocketInMsgTypes.RECEIVE_FA
        'Debug.Print "getFAData "
        getFAData Version
    Case TwsSocketInMsgTypes.HISTORICAL_DATA
        'Debug.Print "getHistoricalData "
        getHistoricalData Version
    Case TwsSocketInMsgTypes.BOND_CONTRACT_DATA
        'Debug.Print "getBondContractData "
        getBondContractData Version
    Case TwsSocketInMsgTypes.SCANNER_PARAMETERS
        'Debug.Print "getScannerParameters "
        getScannerParameters Version
    Case TwsSocketInMsgTypes.SCANNER_DATA
        'Debug.Print "getScannerData "
        getScannerData Version
    Case TwsSocketInMsgTypes.TICK_OPTION_COMPUTATION
        getTickOptionComputation Version
    Case TwsSocketInMsgTypes.TICK_GENERIC
        getTickGeneric Version
    Case TwsSocketInMsgTypes.TICK_STRING
        getTickString Version
    Case TwsSocketInMsgTypes.CONTRACT_DATA_END
        getContractEnd Version
    Case TwsSocketInMsgTypes.OPEN_ORDER_END
        getOpenOrderEnd Version
    Case TwsSocketInMsgTypes.ACCT_DOWNLOAD_END
        getAccountDownloadEnd Version
    Case TwsSocketInMsgTypes.EXECUTION_DATA_END
        getExecutionDataEnd Version
    Case TwsSocketInMsgTypes.DELTA_NEUTRAL_VALIDATION
        getDeltaNeutralValidation Version
    Case TwsSocketInMsgTypes.TICK_SNAPSHOT_END
        getTickSnapshotEnd Version
    Case Else
        s = "Invalid message code from Tws: " & vbCrLf & _
            "msgID=" & msgId & _
            " version=" & Version & _
            " startindex=" & startIndex & vbCrLf & _
            "prevMsgID=" & prevMsgID & _
            " prevVersion=" & prevVersion & _
            " prevStartindex=" & prevStartIndex & vbCrLf & _
            "mIndexIn=" & mInputBufferNextFreeIndex & _
            " mInputIndex=" & mInputParseIndex & vbCrLf & _
            "Buffer contents: " & vbCrLf & _
            formatBuffer
        
        Debug.Print s
        logMessage s, pLogLevel:=LogLevelSevere, pProcName:="checkmessages"
        mTws.ReleaseAPI Me, True
        Exit Sub
    End Select
    
    retryCount = 0
    
    Dim et As Single
    et = mPerformanceElapsedTimer.ElapsedTimeMicroseconds
    With mPerformanceStats(msgId)
        .lastSecondCount = .lastSecondCount + 1
        .lastSecondTime = .lastSecondTime + et
        If et > .longestTime Then .longestTime = et
        If et < .shortestTime Or .shortestTime = 0 Then .shortestTime = et
    End With
    
    eventCount = eventCount + 1
    
    If mElapsedTimer.ElapsedTimeMicroseconds >= 10000000 Then
        logMessage "Event rate per second", pMsgQualifier:=Format(eventCount / 10, "0.0"), pLogLevel:=LogLevelDetail, pProcName:="checkmessages"
        Debug.Print "Event rate per second = " & Format(eventCount / 10, "0.0")
        eventCount = 0
        mElapsedTimer.StartTiming
    End If
    
    checkpointBuffer
    mCheckPointNumber = 1000
    'Debug.Print "Set checkpoint number=" & 1000 & " minputindex=" & mInputParseIndex
Loop

Exit Sub

Err:
Dim errNum As Long
Dim errDesc As String
Dim errSource As String

errNum = Err.number
errDesc = Err.Description
errSource = Err.source

If errNum = DataIncomplete Then
    rollbackBuffer
    retryCount = retryCount + 1
    Debug.Print "Data incomplete count: " & retryCount & " mCheckPointNumber: " & mCheckPointNumber
    Exit Sub
End If

s = "msgID=" & msgId & _
    " version=" & Version & _
    " startindex=" & startIndex & vbCrLf & _
    "prevMsgID=" & prevMsgID & _
    " prevVersion=" & prevVersion & _
    " prevStartindex=" & prevStartIndex & vbCrLf & _
    "mInputBufferNextFreeIndex=" & mInputBufferNextFreeIndex & _
    " mInputParseIndex=" & mInputParseIndex & vbCrLf & _
    "Buffer contents: " & vbCrLf & _
    formatBuffer

Debug.Print "Error " & errNum & ": " & errDesc & vbCrLf & s

logMessage s, pLogLevel:=LogLevelDetail, pProcName:=ProcName
logSocketInputMessage True

gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pErrorNumber:=errNum, pErrorDesc:=errDesc, pErrorSource:=errSource
End Sub

Private Sub checkpointBuffer()
mInputParseIndexChk = mInputParseIndex
End Sub

Private Sub createLoggers()

Const ProcName As String = "createLoggers"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelMediumDetail) Then
    If mPerformanceTimerSecond Is Nothing Then
        Set mPerformanceTimerSecond = CreateIntervalTimer(CDate(Int((Now + 2 * OneSecond) / OneSecond) * OneSecond), _
                                                ExpiryTimeUnitDateTime, _
                                                1000)
        Set mPerformanceTimerPeriod = CreateIntervalTimer(CDate(Int((Now + OneMinute) / OneMinute) * OneMinute), _
                                                ExpiryTimeUnitDateTime, _
                                                60000)
        mPerformanceTimerSecond.StartTimer
        mPerformanceTimerPeriod.StartTimer
        
        Set mPerformanceLogger = GetLogger("tradebuild.log.serviceprovider.ibtwssp.performance")
    End If
End If

Set mSocketLogger = GetLogger("tradebuild.log.serviceprovider.ibtwssp.socket")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub createInMessageBuilder()
Const ProcName As String = "createInMessageBuilder"
On Error GoTo Err

If mInMessageBuilder Is Nothing Then
    Set mInMessageBuilder = CreateStringBuilder
    mInMessageBuilder.Append "IN: "
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub createsocket()
Const ProcName As String = "createsocket"
On Error GoTo Err

If mWinsockForm Is Nothing Then Set mWinsockForm = New WinsockForm
Set mSocket = mWinsockForm.Winsock1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function expiryToDate( _
                ByVal Expiry As String) As Date
Const ProcName As String = "expiryToDate"
On Error GoTo Err

If Len(Expiry) = 8 Then
    expiryToDate = CDate(Left$(Expiry, 4) & "/" & Mid$(Expiry, 5, 2) & "/" & Right$(Expiry, 2))
ElseIf Len(Expiry) = 6 Then
    expiryToDate = CDate(Left$(Expiry, 4) & "/" & Mid$(Expiry, 5, 2) & "/" & "01")
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function formatBuffer() As String
Dim s As StringBuilder
Dim i As Long
Dim j As Long

Const ProcName As String = "formatBuffer"
On Error GoTo Err

Set s = CreateStringBuilder
Do While i < mInputBufferNextFreeIndex
    s.Append Format(i, "0000  ")
    For j = i To i + 49
        If j = mInputBufferNextFreeIndex Then Exit For
        s.Append IIf(mInputBuffer(j) <> 0, Chr$(mInputBuffer(j)), "_")
    Next
    i = i + 50
    If j < mInputBufferNextFreeIndex Then s.AppendLine ""
Loop
formatBuffer = s.ToString

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getBoolean( _
                ByRef fieldName As String) As Boolean
Dim s As String
s = getString(fieldName)
If s = "1" Then
    getBoolean = True
Else
    getBoolean = False
End If
End Function

Private Function getDouble( _
                ByRef fieldName As String) As Double
Dim s As String
s = getString(fieldName)
If s = "" Then
    getDouble = 0#
Else
    getDouble = val(s)
End If
End Function

Private Function getDoubleMax( _
                ByRef fieldName As String) As Double
Dim s As String
s = getString(fieldName)
If s = "" Then
    getDoubleMax = MaxDouble
Else
    getDoubleMax = val(s)
End If
End Function

Private Function getLong( _
                ByRef fieldName As String) As Long
Dim s As String
s = getString(fieldName)
If s = "" Then
    getLong = 0&
Else
    getLong = CLng(s)
End If
End Function

Private Function getLongMax( _
                ByRef fieldName As String) As Long
Dim s As String
s = getString(fieldName)
If s = "" Then
    getLongMax = MaxLong
Else
    getLongMax = CLng(s)
End If
End Function

Private Function getString( _
                ByRef fieldName As String) As String
Dim s As String
Dim i As Long

For i = mInputParseIndex To mInputBufferNextFreeIndex - 1
    If mInputBuffer(i) = 0 Then
        ' we've found our Next string
        'Debug.Print "Got string: mInputIndex=" & mInputParseIndex & " value=" & s
        mInputParseIndex = i + 1
        getString = s
        If Not mInMessageBuilder Is Nothing Then
            mInMessageBuilder.Append fieldName
            mInMessageBuilder.Append "="
            mInMessageBuilder.Append s
            mInMessageBuilder.Append ";"
        End If
        Exit Function
    End If
    ' NB: I've tried various other means of generating the string
    ' without using string concatenation, but for the lengths of the
    ' fields in the TwsAPI protocol (except for error messages) they are all much slower
    ' than string concatenation
    s = s & Chr$(mInputBuffer(i))
Next

' we've got to the end of the current data, so raise an error to abort
' this attempt to assemble a message
'Debug.Print "Failed to get string"
Err.Raise DataIncomplete
End Function

Private Sub getAccountDownloadEnd(ByVal Version As Long)
Dim accountName As String
Const ProcName As String = "getAccountDownloadEnd"

On Error GoTo Err

accountName = getString("Account name")
logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getAccountTime(ByVal Version As Long)
Const ProcName As String = "getAccountTime"
Dim accountTime As String: accountTime = getString("Account Time")


On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getAccountValue(ByVal Version As Long)
Const ProcName As String = "getAccountValue"
Dim Key As String: Key = getString("Key")
Dim val As String: val = getString("Value")
Dim cur As String: cur = getString("Currency")
Dim accountName As String: If Version >= 2 Then accountName = getString("Account Name")


On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getBondContractData(ByVal Version As Long)
' not currently supported
Const ProcName As String = "getBondContractData"
On Error GoTo Err

Dim lRequestId As Long
If (Version >= 3) Then lRequestId = getLong("Req Id")
getString "Symbol"
getString "Sectype"
getString "Cusip"
getDouble "Coupon"
getString "Maturity"
getString "IssueDate"
getString "Ratings"
getString "BondType"
getString "CouponType"
getBoolean "Convertible"
getBoolean "Callable"
getBoolean "Putable"
getString "DescAppend"
getString "Exchange"
getString "Currency"
getString "MarketName"
getString "TradingClass"
getLong "conid"
getDouble "MinTick"
getString "OrderTypes"
getString "ValidExchanges"
If (Version >= 2) Then
    getString "Next Option Date"
    getString "Next Option Type"
    getBoolean "Next Option Partial"
    getString "Notes"
End If
If Version >= 4 Then getString "LongName"


logSocketInputMessage

If Not mTws.ContractDetailsConsumer Is Nothing Then mTws.ContractDetailsConsumer.NotifyError lRequestId, ErrorCodes.ErrUnsupportedOperationException, "Bonds not yet supported"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getContractData(ByVal Version As Long)
Const ProcName As String = "getContractData"
On Error GoTo Err

Dim lReqId As Long
Dim lContract As New TwsContract
Dim lContractDetails As New TwsContractDetails

lContractDetails.Summary = lContract

If (Version >= 3) Then lReqId = getCallersRequestIdFromTwsContractRequestId(getLong("Req Id"))

lContract.Symbol = getString("Symbol")
lContract.Sectype = gTwsSecTypeFromString(getString("Sec type"))
lContract.Expiry = getString("Expiry")
lContract.Strike = getDouble("Strike")
lContract.OptRight = gTwsOptionRightFromString(getString("Right"))
lContract.Exchange = getString("Exchange")
lContract.CurrencyCode = getString("Currency")
lContract.LocalSymbol = getString("Local Symbol")
lContractDetails.MarketName = getString("Market Name")
lContractDetails.TradingClass = getString("Trading class")
lContract.ConId = getLong("Contract id")
lContractDetails.MinTick = getDouble("Minimum tick")
lContract.Multiplier = getLong("Multiplier")
If lContract.Multiplier = 0 Then lContract.Multiplier = 1


lContractDetails.OrderTypes = getString("Order types")
lContractDetails.ValidExchanges = getString("Valid exchanges")
    
If Version >= 2 Then
    lContractDetails.PriceMagnifier = getLong("Price magnifier")
Else
    If lContract.CurrencyCode = "GBP" Then
        lContractDetails.PriceMagnifier = 100
    Else
        lContractDetails.PriceMagnifier = 1
    End If
End If

lContract.Multiplier = lContract.Multiplier / lContractDetails.PriceMagnifier
lContractDetails.MinTick = lContractDetails.MinTick * lContractDetails.PriceMagnifier
    
If Version >= 4 Then lContractDetails.UnderConId = getLong("UnderConId")

If Version >= 5 Then
   lContractDetails.LongName = getString("Long name")
   lContract.PrimaryExch = getString("Primary Exchange")
End If

If Version >= 6 Then
    lContractDetails.ContractMonth = getString("ContractMonth")
    lContractDetails.Industry = getString("Industry")
    lContractDetails.Category = getString("Category")
    lContractDetails.Subcategory = getString("Subcategory")
    lContractDetails.TimeZoneId = getString("TimeZoneId")
    lContractDetails.TradingHours = getString("TradingHours")
    lContractDetails.LiquidHours = getString("LiquidHours")
 End If

logSocketInputMessage

If SerialiseContractDetailsRequests Then
    If Not mGotFirstContractDetails Then
        mGotFirstContractDetails = True
        
        ' now send a reqContract that will give a 'no such security
        ' exists' error. This error will act as a delimiter for all
        ' the contract details returned by the orginal Request. Note that
        ' tests indicate that contract details for different reqContracts
        ' are not interleaved
        
        addMsgId REQ_CONTRACT_DATA
        addString 1, "Version"
        addString "ES", "Symbol"
        addString "FUT", "Sec type"
        addString "200501", "Expiry"
        addString "0", "Strike"
        addString "", "Right"
        addString "LIFFE", "Exchange"
        addString "GBP", "Currency"
        addString "", "Local Symbol"
        send
    End If
End If

If Not mTws.ContractDetailsConsumer Is Nothing Then mTws.ContractDetailsConsumer.NotifyContract lReqId, lContractDetails

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getContractEnd(ByVal Version As Long)
Const ProcName As String = "getContractEnd"
Dim reqId As Long


On Error GoTo Err

reqId = getCallersRequestIdFromTwsContractRequestId(getLong("Req Id"))

logSocketInputMessage

If Not mTws.ContractDetailsConsumer Is Nothing Then mTws.ContractDetailsConsumer.EndContractDetails reqId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getDeltaNeutralValidation(ByVal Version As Long)
Dim reqId As Long: reqId = getLong("reqID")

Dim ConId As Long: ConId = getLong("conID")
Dim Delta As Double: Delta = getDouble("delta")
Dim Price As Double: Price = getDouble("price")

Const ProcName As String = "getDeltaNeutralValidation"

On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getErrorMsg(ByVal Version As Long)
Const ProcName As String = "getErrorMsg"
On Error GoTo Err

Static sLostConnectionToIb As Boolean
Dim id As Long
Dim lErrorCode As Long
Dim lErrorMsg As String
Dim execFilter As TwsExecutionFilter
Dim index As Long
Dim mapEntry As OrderIdMapEntry


id = getLong("Id")
Debug.Print "Error id: " & id

lErrorCode = getLong("Error code")
Debug.Print "Error code: " & lErrorCode

lErrorMsg = getString("Error msg")
Debug.Print "Error msg: " & lErrorMsg

logSocketInputMessage True

Select Case getIdType(id)
Case IdTypeContractData
    processContractError getCallersRequestIdFromTwsContractRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeRealtimeData
    If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyError getCallersRequestIdFromTwsMarketDataRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeMarketDepth
    If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.NotifyError getCallersRequestIdFromTwsMarketDepthRequestId(id), lErrorCode, lErrorMsg
    Exit Sub
Case IdTypeHistoricalData
    If lErrorCode = 165 Then
        logMessage "Connected to IB Historical Market Data Service - " & lErrorMsg, pProcName:=ProcName
    Else
        If Not mTws.HistDataConsumer Is Nothing Then mTws.HistDataConsumer.NotifyError getCallersRequestIdFromTwsHistRequestId(id), lErrorCode, lErrorMsg
    End If
    Exit Sub
Case IdTypeOrder
    If sLostConnectionToIb And lErrorCode = 200 Then Exit Sub
    If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.NotifyError id, lErrorCode, lErrorMsg
    Exit Sub
End Select

Select Case lErrorCode
Case 316
    If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.ResetMarketDepth pReEstablish:=True
Case 317
    If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.ResetMarketDepth pReEstablish:=False
Case 321    ' invalid Request
    Select Case getIdType(id)
    Case IdTypeNone
        logMessage "Error (" & lErrorCode & "; id=" & id & ") from Tws: " & lErrorMsg, pProcName:=ProcName
        'mCommonServiceConsumer.NotifyRecoverableError lErrorCode, lErrorMsg, -1
    Case IdTypeContractData
    Case IdTypeRealtimeData
    Case IdTypeMarketDepth
    Case IdTypeHistoricalData
    Case IdTypeOrder
    End Select
Case 326    ' ClientID already in use

Case 1100, 2110
    ' connectivity between Tws and IB has been lost
    
    sLostConnectionToIb = True
    RaiseEvent ConnectionToIBClosed
    logMessage "Connection to IB has been lost", pProcName:=ProcName
    
Case 1101
    ' connectivity between Tws and IB has been restored, but data has been
    ' lost, so need to reestablish market data and market depth requests

    sLostConnectionToIb = False
    
    reEstablishMarketData False
    If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.ResetMarketDepth True
    
    ' Now need to reassociate Order ids with tws
    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New TwsExecutionFilter
    execFilter.ClientID = mClientID
    RequestExecutions 1, execFilter

    RaiseEvent ConnectionToIBRecovered
    
    logMessage "Connection to IB recovered: market data re-established", pProcName:=ProcName
    
Case 1102
    ' connectivity between Tws and IB has been restored without loss of data
    ' Now need to reassociate Order ids with tws

    sLostConnectionToIb = False
    
    RequestOpenOrders
    
    ' Now make sure we're aware of any executions that occurred during
    ' the disconnection
    Set execFilter = New TwsExecutionFilter
    execFilter.ClientID = mClientID
    RequestExecutions 1, execFilter

    RaiseEvent ConnectionToIBRecovered
    
    logMessage "Connection to IB recovered: no loss of data", pProcName:=ProcName
    
Case 2103, 2104, 2105, 2106, 2107, 2108
    logMessage lErrorMsg, pProcName:=ProcName
Case Else
    Select Case getIdType(id)
    Case IdTypeOrder
    Case Else
        logMessage "Error (" & lErrorCode & "; id=" & id & ") from Tws: " & lErrorMsg, pProcName:=ProcName
        'mCommonServiceConsumer.NotifyRecoverableError lErrorCode, lErrorMsg, -1
    End Select
    
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getExecutionData(ByVal Version As Long)
Const ProcName As String = "getExecutionData"
On Error GoTo Err

Dim lReqId As Long

Dim lExecution As New TwsExecution
Dim lContract As New TwsContract

lReqId = -1
If (Version >= 7) Then lReqId = getCallersRequestIdFromTwsExecutionsRequestId(getLong("lReqId"))

lExecution.OrderId = getLong("OrderId")

With lContract
    If (Version >= 5) Then .ConId = getLong("Con Id")
    .Symbol = getString("Symbol")
    .Sectype = gTwsSecTypeFromString(getString("Sec type"))
    .Expiry = getString("Expiry")
    .Strike = getDouble("Strike")
    .OptRight = gTwsOptionRightFromString(getString("Right"))
    .Exchange = getString("Exchange")
    .CurrencyCode = getString("Currency")
    .LocalSymbol = getString("Local Symbol")
End With

With lExecution
    .ExecId = getString("Exec id")
    .Time = getString("Exec Time")
    .AcctNumber = getString("Account id")
    .Exchange = getString("Exchange")
    .Side = getString("Side")
    .Shares = getLong("Shares")
    .Price = getDouble("Price")
    .Price = adjustPrice(.Price, lContract.CurrencyCode)
        
    If Version >= 2 Then .PermId = getLong("Perm Id")
    
    If Version >= 3 Then .ClientID = getLong("Client id")
       
    If Version >= 4 Then .Liquidation = getBoolean("Liquidation")
    
    If (Version >= 6) Then
        .CumQty = getLong("Cum Qty")
        .AvgPrice = getDouble("Avg Price")
    End If
End With

logSocketInputMessage True

If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.NotifyExecution lReqId, lContract, lExecution

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getExecutionDataEnd(ByVal Version As Long)
Const ProcName As String = "getExecutionDataEnd"
On Error GoTo Err

Dim reqId  As Long

reqId = getCallersRequestIdFromTwsExecutionsRequestId(getLong("reqId"))
logSocketInputMessage

If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.EndExecutions reqId
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getFAData(ByVal Version As Long)
Const ProcName As String = "getFAData"
Dim DataType As FADataTypes: DataType = getLong("Data type")
Dim XMLData As String: XMLData = getString("XML data")


On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getHistoricalData(ByVal Version As Long)
Const ProcName As String = "getHistoricalData"

On Error GoTo Err

Static sCheckPointNumber As Long
Static sStartInfo As TWSHistDataStartInfo
Static sBar As TWSBar
Static i As Long

Dim lBar As TWSBar

If sCheckPointNumber = 0 Then
    i = 0
    sBar = lBar
    
    sStartInfo.RequestId = getCallersRequestIdFromTwsHistRequestId(getLong("Request id"))
    
    checkpointBuffer
    sCheckPointNumber = 100
End If

If sCheckPointNumber = 100 Then
    If Version >= 2 Then
        sStartInfo.StartDate = getString("Start date")      ' startDateStr
    End If
    checkpointBuffer
    sCheckPointNumber = 110
End If

If sCheckPointNumber = 110 Then
    If Version >= 2 Then
        sStartInfo.EndDate = getString("End date")      ' endDateStr
    End If
    checkpointBuffer
    sCheckPointNumber = 120
End If

If sCheckPointNumber = 120 Then
    sStartInfo.BarCount = getLong("Item count")
    Debug.Print "Bars retrieved: " & sStartInfo.BarCount
    
    If Not mTws.HistDataConsumer Is Nothing Then mTws.HistDataConsumer.StartHistData sStartInfo
    
    checkpointBuffer
    sCheckPointNumber = 200
End If

If sCheckPointNumber >= 200 Then
    For i = i To sStartInfo.BarCount - 1
        If sCheckPointNumber = 200 Then
            sBar.Timestamp = getString("Bar date")
            
            checkpointBuffer
            sCheckPointNumber = 210
        End If
        If sCheckPointNumber = 210 Then
            sBar.OpenValue = getDouble("Open")
            checkpointBuffer
            sCheckPointNumber = 220
        End If
        If sCheckPointNumber = 220 Then
            sBar.HighValue = getDouble("High")
            checkpointBuffer
            sCheckPointNumber = 230
        End If
        If sCheckPointNumber = 230 Then
            sBar.LowValue = getDouble("Low")
            checkpointBuffer
            sCheckPointNumber = 240
        End If
        If sCheckPointNumber = 240 Then
            sBar.CloseValue = getDouble("Close")
            checkpointBuffer
            sCheckPointNumber = 250
        End If
        If sCheckPointNumber = 250 Then
            sBar.Volume = getLong("Volume")
            checkpointBuffer
            sCheckPointNumber = 260
        End If
        If sCheckPointNumber = 260 Then
            sBar.WAP = getDouble("WAP")
            checkpointBuffer
            sCheckPointNumber = 270
        End If
        If sCheckPointNumber = 270 Then
            sBar.HasGaps = getBoolean("Has gaps")
            checkpointBuffer
            sCheckPointNumber = 280
        End If
        If sCheckPointNumber = 280 Then
            If Version >= 3 Then
                sBar.TickVolume = getLong("Tick volume")
                checkpointBuffer
                sCheckPointNumber = 280
            End If
        End If
        
        If Not mTws.HistDataConsumer Is Nothing Then mTws.HistDataConsumer.NotifyBar sBar
        
        sCheckPointNumber = 200
    Next
End If
sCheckPointNumber = 0

logSocketInputMessage

If Not mTws.HistDataConsumer Is Nothing Then mTws.HistDataConsumer.EndHistData

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getCallersRequestIdFromTwsContractRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsContractRequestId = pId - BaseContractRequestId
End Function

Private Function getCallersRequestIdFromTwsExecutionsRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsExecutionsRequestId = pId - BaseExecutionsRequestId
End Function

Private Function getCallersRequestIdFromTwsHistRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsHistRequestId = pId - BaseHistoricalDataRequestId
End Function

Private Function getCallersRequestIdFromTwsMarketDataRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsMarketDataRequestId = pId - BaseMarketDataRequestId
End Function

Private Function getCallersRequestIdFromTwsMarketDepthRequestId(ByVal pId As Long) As Long
getCallersRequestIdFromTwsMarketDepthRequestId = pId - BaseMarketDepthRequestId
End Function

Private Function getTwsContractRequestIdFromCallersRequestId(ByVal pId As Long) As Long
If pId > MaxCallersContractRequestId Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Max request id is " & MaxCallersContractRequestId
getTwsContractRequestIdFromCallersRequestId = pId + BaseContractRequestId
End Function

Private Function getTwsExecutionsRequestIdFromCallersRequestId(ByVal pId As Long) As Long
If pId > MaxCallersExecutionsRequestId Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Max request id is " & MaxCallersExecutionsRequestId
getTwsExecutionsRequestIdFromCallersRequestId = pId + BaseExecutionsRequestId
End Function

Private Function getTwsHistRequestIdFromCallersRequestId(ByVal pId As Long) As Long
If pId > MaxCallersHistoricalDataRequestId Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Max request id is " & MaxCallersHistoricalDataRequestId
getTwsHistRequestIdFromCallersRequestId = pId + BaseHistoricalDataRequestId
End Function

Private Function getTwsMarketDataRequestIdFromCallersRequestId(ByVal pId As Long) As Long
If pId > MaxCallersMarketDataRequestId Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Max request id is " & MaxCallersMarketDataRequestId
getTwsMarketDataRequestIdFromCallersRequestId = pId + BaseMarketDataRequestId
End Function

Private Function getTwsMarketDepthRequestIdFromCallersRequestId(ByVal pId As Long) As Long
If pId > MaxCallersMarketDepthRequestId Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Max request id is " & MaxCallersMarketDepthRequestId
getTwsMarketDepthRequestIdFromCallersRequestId = pId + BaseMarketDepthRequestId
End Function

Private Function getIdType( _
                ByVal id As Long) As IdTypes
Const ProcName As String = "getIdType"
On Error GoTo Err

If id >= BaseOrderId Then
    getIdType = IdTypeOrder
ElseIf id >= BaseContractRequestId Then
    getIdType = IdTypeContractData
ElseIf id >= BaseExecutionsRequestId Then
    getIdType = IdTypeExecution
ElseIf id >= BaseHistoricalDataRequestId Then
    getIdType = IdTypeHistoricalData
ElseIf id >= BaseMarketDepthRequestId Then
    getIdType = IdTypeMarketDepth
ElseIf id >= 0 Then
    getIdType = IdTypeRealtimeData
Else
    getIdType = IdTypeNone
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub getManagedAccounts(ByVal Version As Long)
Const ProcName As String = "getManagedAccounts"
Dim accountsList As String: accountsList = getString("Accounts list")


On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getMarketDepth(ByVal Version As Long)
Const ProcName As String = "getMarketDepth"
On Error GoTo Err

Dim id As Long: id = getLong("Id")
Dim lPosition As Long: lPosition = getLong("Position")
Dim lOperation As DOMOperations:  lOperation = getLong("Operation")
Dim lSide As DOMSides: lSide = getLong("Side")
Dim lPrice As Double: lPrice = getDouble("Price")
Dim lSize As Long: lSize = getLong("Size")

logSocketInputMessage

If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.NotifyMarketDepth getCallersRequestIdFromTwsMarketDepthRequestId(id), lPosition, "", lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getMarketDepthL2(ByVal Version As Long)
Const ProcName As String = "getMarketDepthL2"
On Error GoTo Err

Dim id As Long: id = getLong("Id")
Dim lPosition As Long: lPosition = getLong("Position")
Dim lMarketMaker As String: lMarketMaker = getString("Marketmaker")
Dim lOperation As DOMOperations:  lOperation = getLong("Operation")
Dim lSide As DOMSides: lSide = getLong("Side")
Dim lPrice As Double: lPrice = getDouble("Price")
Dim lSize As Long: lSize = getLong("Size")

logSocketInputMessage

If Not mTws.MarketDepthConsumer Is Nothing Then mTws.MarketDepthConsumer.NotifyMarketDepth getCallersRequestIdFromTwsMarketDepthRequestId(id), lPosition, lMarketMaker, lOperation, lSide, lPrice, lSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getNewsBulletins(ByVal Version As Long)
Const ProcName As String = "getNewsBulletins"
Dim msgId As Long: msgId = getLong("Id")
Dim msgType As Long: msgType = getLong("Type")
Dim newsMessage As String: newsMessage = getString("Message")
Dim originatingExch As String: originatingExch = getString("Orig Exchange")


On Error GoTo Err

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getNextValidId(ByVal Version As Long)
Const ProcName As String = "getNextValidId"

On Error GoTo Err

mNextOrderID = getLong("Next id")

logSocketInputMessage

If mNextOrderID < BaseOrderId Then mNextOrderID = BaseOrderId

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getOpenOrder(ByVal Version As Long)
Const ProcName As String = "getOpenOrder"
On Error GoTo Err

Dim lContract As TwsContract
Dim lOrder As TwsOrder
Dim lOrderState As TwsOrderState


Set lOrder = New TwsOrder
Set lContract = New TwsContract
Set lOrderState = New TwsOrderState

lOrder.OrderId = getLong("Id")

' read contract fields
With lContract
    If Version >= 17 Then lContract.ConId = getLong("conId")
    .Symbol = getString("Symbol")
    .Sectype = gTwsSecTypeFromString(getString("Sec type"))
    .Expiry = getString("Expiry")
    .Strike = getDouble("Strike")
    .OptRight = gTwsOptionRightFromString(getString("Right"))
    .Exchange = getString("Exchange")
    .CurrencyCode = getString("Currency")
    If Version >= 2 Then .LocalSymbol = getString("Local Symbol")
End With

' read Order fields
With lOrder
    .Action = gTwsOrderActionFromString(getString("Action"))
    .TotalQuantity = getLong("Quantity")
    .OrderType = gTwsOrderTypeFromString(getString("Order type"))
    .LmtPrice = getDouble("Limit price")
    .AuxPrice = getDouble("Aux price")
    .Tif = gTwsOrderTIFFromString(getString("Time in force"))
    .OcaGroup = getString("OCA group")
    .Account = getString("Account")
    .OpenClose = getString("Open/close")
    .Origin = getLong("Origin")
    .OrderRef = getString("Order ref")
    If Version >= 3 Then .ClientID = getLong("Client id")
    
    If Version >= 4 Then
        .PermId = getLong("Perm id")
        .OutsideRth = getBoolean("Outside RTH")
        .Hidden = getBoolean("Hidden")
        .DiscretionaryAmt = getDouble("Discr amt")
    End If

    If Version >= 5 Then .GoodAfterTime = getString("Good after Time")
    
    If Version >= 6 Then getString "Shares allocation"
    
    If Version >= 7 Then
        .FaGroup = getString("FA Group")
        .FaMethod = getString("FA method")
        .FaPercentage = getString("FA Percentage")
        .FaProfile = getString("FA Profile")
    End If
    
    If Version >= 8 Then .GoodAfterTime = getString("Good till date")
    
    If Version >= 9 Then
        .Rule80A = getString("Rule80A")
        .PercentOffset = getDouble("Percent Offset")
        .SettlingFirm = getString("Settling Firm")
        .ShortSaleSlot = getLong("shortSaleSlot")
        .DesignatedLocation = getString("designatedLocation")
        .AuctionStrategy = getLong("Auction Strategy")
        .StartingPrice = getDouble("starting Price")
        .StockRefPrice = getDouble("stockRefPrice")
        .Delta = getDouble("delta")
        .StockRangeLower = getDouble("stockRangeLower")
        .StockRangeUpper = getDouble("stockRangeUpper")
        .DisplaySize = getLong("displaySize")
        'setproperty .providerProperties, "rthOnly"
        .BlockOrder = getBoolean("blockOrder")
        .SweepToFill = getBoolean("sweepToFill")
        .AllOrNone = getBoolean("allOrNone")
        .MinQty = getLong("minQty")
        .OcaType = getLong("ocaType")
        .ETradeOnly = getBoolean("eTradeOnly")
        .FirmQuoteOnly = getBoolean("firmQuoteOnly")
        .NbboPriceCap = getDouble("nbboPriceCap")
    End If

    If Version >= 10 Then
        .ParentId = getLong("ParentId")
        .TriggerMethod = getLong("triggerMethod")
    End If

    If Version >= 11 Then
        .Volatility = getDouble("volatility")
        .VolatilityType = getLong("volatilityType")
        If Version = 11 Then
            .DeltaNeutralOrderType = IIf(getLong("deltaNeutralOrderType") = 0, "NONE", "MKT")
        Else ' version 12 and up
            .DeltaNeutralOrderType = getString("deltaNeutralOrderType")
            .DeltaNeutralAuxPrice = getDouble("deltaNeutralAuxPrice")
        End If
        .ContinuousUpdate = getLong("continuousUpdate")
        If mServerVersion = 26 Then
            .StockRangeLower = getDouble("stockRangeLower")
            .StockRangeUpper = getDouble("stockRangeUpper")
        End If
        .ReferencePriceType = getLong("referencePriceType")
    End If
    
    If Version >= 13 Then .TrailStopPrice = getDouble("trailStopPrice")

    If Version >= 14 Then
        .BasisPoints = getDouble("basisPoints")
        .BasisPointsType = getLong("basisPointsType")
        lContract.ComboLegsDescription = getString("comboLegsDescrip")
    End If

    If (Version >= 15) Then
        If (Version >= 20) Then
            .ScaleInitLevelSize = getLongMax("Scale Init Level Size")
            .ScaleSubsLevelSize = getLongMax("Scale Subs Level Size")
        Else
            .ScaleInitLevelSize = getLongMax("Scale Init Level Size")
        End If
        .ScalePriceIncrement = getDoubleMax("Scale Price Increment")
    End If
    
    If (Version >= 19) Then
        .ClearingAccount = getString("Clearing Account")
        .ClearingIntent = getString("Clearing Intent")
    End If
    
    If Version >= 22 Then .NotHeld = getBoolean("Not held")
    
    If (Version >= 20) Then
        Dim lUnderComp As TwsUnderComp
        If CBool(getLong("UnderComp")) Then
            lUnderComp.ConId = getLong("UnderComp ConId")
            lUnderComp.Delta = getDouble("UnderComp Delta")
            lUnderComp.Price = getDouble("UnderComp Price")
            lContract.UnderComp = lUnderComp
        End If
    End If
    
    If Version >= 21 Then
        .AlgoStrategy = getString("Algo strategy")
        If .AlgoStrategy <> "" Then
            Dim algoParamsCount As Long
            algoParamsCount = getLong("Algo params count")
            If algoParamsCount > 0 Then
                ReDim lAlgoParams(algoParamsCount - 1) As TwsTagValue
                Dim i As Long
                For i = 0 To algoParamsCount - 1
                    lAlgoParams(i).Tag = getString("Tag")
                    lAlgoParams(i).Value = getString("Value")
                Next
                .AlgoParams = lAlgoParams
            End If
        End If
    End If
    
    If (Version >= 16) Then

        .WhatIf = getBoolean("What If")

        With lOrderState
            .Status = getString("Status")
            .InitMargin = getString("Init Margin")
            .MaintMargin = getString("Maint Margin")
            .EquityWithLoan = getString("Equity With Loan")
            .Commission = getDoubleMax("Commission")
            .MinCommission = getDoubleMax("Min Commission")
            .MaxCommission = getDoubleMax("Max Commission")
            .CommissionCurrency = getString("Commission Currency")
            .WarningText = getString("Warning Text")
        End With
    End If
End With

logSocketInputMessage True

If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.NotifyOpenOrder lOrder.OrderId, lContract, lOrder, lOrderState

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getOpenOrderEnd(ByVal Version As Long)
Const ProcName As String = "getOpenOrderEnd"

On Error GoTo Err

logSocketInputMessage

If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.EndOpenOrders
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getOrderStatus(ByVal Version As Long)
Const ProcName As String = "getOrderStatus"
On Error GoTo Err

' NB: lAvgFillPrice and lLastFillPrice ought to be subjected to adjustPrice,
' but we can't call that here because we don't know the currency code. But not really
' an issue since we are not likely to encounter a TWS with server version < 23 !

Dim lorderId As Long: lorderId = getLong("id")
Dim lStatus As String: lStatus = getString("Status")
Dim lFilled As Long: lFilled = getLong("Filled")
Dim lRemaining As Long: lRemaining = getLong("Remaining")
Dim lAvgFillPrice As Double:  lAvgFillPrice = getDouble("Avg fill price")
Dim lLastFillPrice As Double
Dim lClientID As Long
Dim lPermId As Long
Dim lParentId As Long
Dim lWhyHeld As String

If Version >= 2 Then lPermId = getLong("Perm id")

If Version >= 3 Then lParentId = getLong("Parent id")

If Version >= 4 Then lLastFillPrice = getDouble("Last fill price")

If Version >= 5 Then lClientID = getLong("Client id")

If Version >= 6 Then lWhyHeld = getString("Why held")

logSocketInputMessage True

If Not mTws.OrderInfoConsumer Is Nothing Then mTws.OrderInfoConsumer.NotifyOrderStatus lorderId, lStatus, lFilled, lRemaining, lAvgFillPrice, lPermId, lParentId, lLastFillPrice, lClientID, lWhyHeld

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getPortfolioValue(ByVal Version As Long)
Const ProcName As String = "getPortfolioValue"
On Error GoTo Err

Dim Symbol As String
Dim LocalSymbol As String
Dim Sectype As String ' keep the socket string for logging
Dim Expiry As String
Dim Strike As Double
Dim rightString As String
Dim Exchange As String
Dim CurrencyCode As String

Dim lContractSpecifier As ContractSpecifier

If Version >= 6 Then getLong "Contract id"

Symbol = getString("Symbol")
Sectype = getString("Sec type")
Expiry = getString("Expiry")
Strike = getDouble("Strike")
rightString = getString("Right")

If Version >= 7 Then
    getString "Multiplier"
    getString "Primary exchg"
End If

CurrencyCode = getString("Currency")

If Version >= 2 Then LocalSymbol = getString("Locale Symbol")

Set lContractSpecifier = CreateContractSpecifier(LocalSymbol, _
                                                Symbol, _
                                                Exchange, _
                                                SecTypeFromString(Sectype), _
                                                CurrencyCode, _
                                                Expiry, _
                                                Strike, _
                                                OptionRightFromString(rightString))

Dim position As Long: position = getLong("Position")
Dim marketPrice As Double: marketPrice = getDouble("Market price")
Dim marketValue As Double: marketValue = getDouble("Market Value")
Dim averageCost As Double: averageCost = 0#
Dim unrealizedPNL As Double: unrealizedPNL = 0#
Dim realizedPNL As Double: realizedPNL = 0#
If Version >= 3 Then
    averageCost = getDouble("Avg cost")
    unrealizedPNL = getDouble("Unrealized PNL")
    realizedPNL = getDouble("Realized PNL")
End If

Dim accountName As String: If Version >= 4 Then accountName = getString("Account Name")

If Version = 6 And mServerVersion = 39 Then getString "Primary Exchange"

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getScannerData(ByVal Version As Long)
Const ProcName As String = "getScannerData"
On Error GoTo Err

Dim lContractSpecifier As ContractSpecifier
Dim lContractBuilder As ContractBuilder
Dim TickerId As Long
Dim numberOfElements As Long
Dim rank As Long
Dim distance As String
Dim benchmark As String
Dim projection As String
Dim i As Long

Dim Symbol As String
Dim LocalSymbol As String
Dim Sectype As SecurityTypes
Dim Expiry As String
Dim Strike As Double
Dim OptRight As OptionRights
Dim Exchange As String
Dim CurrencyCode As String

TickerId = getLong("Id")
numberOfElements = getLong("Number of elements")

For i = 0 To numberOfElements - 1
    rank = getLong("Rank")

    If (Version >= 3) Then getLong "ConId"
    
    Symbol = getString("Symbol")
    Sectype = SecTypeFromString(getString("Sec type"))
    Expiry = getString("Expiry")
    Strike = getDouble("Stirke")
    OptRight = OptionRightFromString(getString("Right"))
    Exchange = getString("Exchange")
    CurrencyCode = getString("Currency")
    LocalSymbol = getString("Local Symbol")
    
    Set lContractSpecifier = CreateContractSpecifier(LocalSymbol, _
                                                    Symbol, _
                                                    Exchange, _
                                                    Sectype, _
                                                    CurrencyCode, _
                                                    Expiry, _
                                                    Strike, _
                                                    OptRight)
    Set lContractBuilder = CreateContractBuilder(lContractSpecifier)
    getString "marketName"
    getString "tradingClass"
    distance = getString("Distance")
    benchmark = getString("Benchmark")
    projection = getString("Projection")
    If (Version >= 2) Then getString "Legs Str "
Next

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getScannerParameters(ByVal Version As Long)
Const ProcName As String = "getScannerParameters"
Dim xml  As String


On Error GoTo Err

xml = getString("XML")

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getTickGeneric(ByVal Version As Long)
Const ProcName As String = "getTickGeneric"
On Error GoTo Err

Dim lTickerId As Long
Dim lTickType As Long
Dim lValue As Double

lTickerId = getLong("tickerId")
lTickType = getLong("tickType")
lValue = getDouble("Value")

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyTickGeneric getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lValue

logSocketInputMessage

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getTickOptionComputation(ByVal Version As Long)
Dim lTickerId As Long: lTickerId = getLong("tickerId")
Dim lTickType As Long: lTickType = getLong("tickType")
Dim lImpliedVol As Double: lImpliedVol = getDouble("impliedVol")

Const ProcName As String = "getTickOptionComputation"

On Error GoTo Err

If lImpliedVol < 0 Then
    ' -1 is the "not yet computed" indicator
    lImpliedVol = MaxDouble
End If

Dim lDelta As Double: lDelta = getDouble("delta")
If Abs(lDelta) > 1 Then
    ' -2 is the "not yet computed" indicator
    lDelta = MaxDouble
End If

Dim lOptPrice As Double: lOptPrice = MaxDouble
Dim lPvDividend As Double: lPvDividend = MaxDouble
Dim lGamma As Double: lGamma = MaxDouble
Dim lVega As Double: lVega = MaxDouble
Dim lTheta As Double: lTheta = MaxDouble
Dim lUndPrice As Double: lUndPrice = MaxDouble

If Version >= 6 Or lTickType = TwsTickTypes.TwsTickModelOption Then   ' introduced in version == 5
    lOptPrice = getDouble("optPrice")
    If lOptPrice < 0 Then   ' -1 is the "not yet computed" indicator
        lOptPrice = MaxDouble
    End If
    lPvDividend = getDouble("pvDividend")
    If lPvDividend < 0 Then    ' -1 is the "not yet computed" indicator
        lPvDividend = MaxDouble
    End If
End If

If (Version >= 6) Then
    lGamma = getDouble("gamma")
    If (Abs(lGamma) > 1) Then ' -2 is the "not yet computed" indicator
        lGamma = MaxDouble
    End If
    lVega = getDouble("Vega")
    If (Abs(lVega) > 1) Then ' -2 is the "not yet computed" indicator
        lVega = MaxDouble
    End If
    lTheta = getDouble("theta")
    If (Abs(lTheta) > 1) Then ' -2 is the "not yet computed" indicator
        lTheta = MaxDouble
    End If
    lUndPrice = getDouble("undPrice")
    If (lUndPrice < 0) Then ' -1 is the "not yet computed" indicator
        lUndPrice = MaxDouble
    End If
End If

logSocketInputMessage

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyTickOptionComputation getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lImpliedVol, lDelta, lOptPrice, lPvDividend, lGamma, lVega, lTheta, lUndPrice

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getTickPrice(ByVal Version As Long)
Const ProcName As String = "getTickPrice"
On Error GoTo Err

Dim lTickType As Long
Dim lTickerId As Long
Dim lPrice As Double
Dim lSize As Long
Dim lCanAutoExecute As Boolean

lTickerId = getLong("Ticker id")
lTickType = getLong("Tick type")
lPrice = getDouble("Price")

If Version >= 2 Then lSize = getLong("Size")
If Version >= 3 Then lCanAutoExecute = getBoolean("Can autoexecute")

logSocketInputMessage

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyTickPrice getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lPrice, lSize, lCanAutoExecute

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getTickSize(ByVal Version As Long)
Const ProcName As String = "getTickSize"
Dim lTickType As Long
Dim lTickerId As Long
Dim lSize As Long

On Error GoTo Err

lTickerId = getLong("Ticker id")
lTickType = getLong("Tick type")
lSize = getLong("Size")

logSocketInputMessage

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyTickSize getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub getTickSnapshotEnd(ByVal Version As Long)
Const ProcName As String = "getTickSnapshotEnd"
On Error GoTo Err

Dim lReqId As Long: lReqId = getLong("reqID")
logSocketInputMessage

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.EndTickSnapshot getCallersRequestIdFromTwsMarketDataRequestId(lReqId)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub getTickString(ByVal Version As Long)
Const ProcName As String = "getTickString"
On Error GoTo Err

Dim lTickerId As Long: lTickerId = getLong("tickerId")
Dim lTickType As Long: lTickType = getLong("tickType")
Dim lValue As String: lValue = getString("Value")

logSocketInputMessage

If Not mTws.MarketDataConsumer Is Nothing Then mTws.MarketDataConsumer.NotifyTickString getCallersRequestIdFromTwsMarketDataRequestId(lTickerId), lTickType, lValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub handleTwsDisconnection()
Const ProcName As String = "handleTwsDisconnection"
On Error GoTo Err

mConnectionState = ConnectionStates.ConnNotConnected
Set mSocket = Nothing

If mDisconnectedByApplication Then
    logMessage "Connection to Tws closed: " & ConnectionString, pProcName:=ProcName
Else
    If mConnectionRetryIntervalSecs = 0 Then
        RaiseEvent ConnectionToTwsClosed(False)
        logMessage "Connection to Tws closed" & _
                    ConnectionString, _
                    ProcName
    Else
        retryConnection
        RaiseEvent ConnectionToTwsClosed(True)
        logMessage "Connection to Tws closed - reconnecting: " & _
                    ConnectionString, _
                    ProcName
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function legOpenCloseFromString(ByVal Value As String) As TwsLegOpenCloseCodes
Select Case UCase$(Value)
Case "0"
    legOpenCloseFromString = LegSamePos
Case "1"
    legOpenCloseFromString = LegOpenPos
Case "2"
    legOpenCloseFromString = LegClosePos
Case "3"
    legOpenCloseFromString = LegUnknownPos
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid leg open/close value"
End Select
End Function

Private Function legOpenCloseToString(ByVal Value As TwsLegOpenCloseCodes) As String
Select Case Value
Case TwsLegOpenCloseCodes.TwsLegOpenCloseSame, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseOpen, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseClose, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseUnknown
    legOpenCloseToString = CStr(Value)
Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid leg open/close value"
End Select
End Function

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub logSocketInputMessage(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "logSocketInputMessage"
On Error GoTo Err

Static prevMessageBuilder As StringBuilder

If ignoreLogLevel Then
    If Not mInMessageBuilder Is Nothing Then
        mSocketLogger.Log LogLevelNormal, mInMessageBuilder.ToString
        Set prevMessageBuilder = mInMessageBuilder
    ElseIf Not prevMessageBuilder Is Nothing Then
        mSocketLogger.Log LogLevelNormal, prevMessageBuilder.ToString
        Set prevMessageBuilder = Nothing
    End If
ElseIf mSocketLogger.IsLoggable(LogLevelHighDetail) Then
    mSocketLogger.Log LogLevelHighDetail, mInMessageBuilder.ToString
    Set prevMessageBuilder = mInMessageBuilder
Else
    Set prevMessageBuilder = mInMessageBuilder
End If
Set mInMessageBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub logSocketOutputMessage(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "logSocketOutputMessage"
On Error GoTo Err

If ignoreLogLevel Then
    mSocketLogger.Log LogLevelNormal, mOutMessageBuilder.ToString
ElseIf mSocketLogger.IsLoggable(LogLevelHighDetail) Then
    mSocketLogger.Log LogLevelHighDetail, mOutMessageBuilder.ToString
End If
Set mOutMessageBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function OptionRightFromString(ByVal Value As String) As OptionRights
Select Case UCase$(Value)
Case "C"
    OptionRightFromString = OptCall
Case "P"
    OptionRightFromString = OptPut
Case Else
    OptionRightFromString = OptNone
End Select
End Function

Private Function OptionRightToString(ByVal Value As OptionRights) As String
Select Case Value
Case OptCall
    OptionRightToString = "CALL"
Case OptPut
    OptionRightToString = "PUT"
End Select
End Function

Private Sub processContractError(ByVal pRequestId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "processContractError"
On Error GoTo Err

If SerialiseContractDetailsRequests And mGotFirstContractDetails Then
    ' we are processing the forced invalid contract details Request
    ' for a Request for which we have received one or more contract details
    mContractDetailsRequestInProgress = False
    If Not mTws.ContractDetailsConsumer Is Nothing Then mTws.ContractDetailsConsumer.EndContractDetails pRequestId
Else
    ' this is a genuine error response to a contract details Request
    If Not mTws.ContractDetailsConsumer Is Nothing Then mTws.ContractDetailsConsumer.NotifyError pRequestId, pErrorCode, pErrorMsg
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reEstablishMarketData(ByVal newConnection As Boolean)
'Const ProcName As String = "reEstablishMarketData"
'Dim i As Long
'
'On Error GoTo Err
'
'logMessage "Re-establishing market data: newConnection = " & CStr(newConnection), ProcName
'For i = 0 To UBound(mTickerTable)
'    With mTickerTable(i)
'        If Not .Contract Is Nothing Then
'            If Not newConnection Then cancelMarketDataEx i, False
'            .Incarnation = .Incarnation + IncarnationIncrement
'            reqMarketData i
'        End If
'    End With
'    clearPrevPrices i
'Next
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub ReleaseSocket()
Const ProcName As String = "ReleaseSocket"

On Error GoTo Err

If Not mSocket Is Nothing Then
    mSocket.Close
    Set mSocket = Nothing
    Unload mWinsockForm
    Set mWinsockForm = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub reqContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "reqContractDetails"
On Error GoTo Err

Const Version = 5

If mConnectionState <> ConnectionStates.ConnConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"

If pContract.Sectype = TwsSecTypeCombo Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Combo contracts are not supported"

If SerialiseContractDetailsRequests Then
    If mContractDetailsRequestInProgress Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Contract details requests must be serialised with this version of TWS"
    mContractDetailsRequestInProgress = True
    mGotFirstContractDetails = False
End If

addMsgId REQ_CONTRACT_DATA
addString Version, "Version"

If mServerVersion >= MIN_SERVER_VER_CONTRACT_DATA_CHAIN Then
    addString getTwsContractRequestIdFromCallersRequestId(pRequestId), "Request id"
End If

With pContract
    If mServerVersion >= MIN_SERVER_VER_CONTRACT_CONID Then addString .ConId, "Contract id"
    addString .Symbol, "Symbol"
    addString gTwsSecTypeToShortString(.Sectype), "Sec type"
    addString .Expiry, "Expiry"
    addDouble .Strike, "Strike"
    addString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then addString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    addString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    addString .CurrencyCode, "Currency"
    If (.Sectype = SecTypeFuture Or _
        .Sectype = SecTypeFuturesOption Or _
        .Sectype = SecTypeOption) _
    Then
        If expiryToDate(.Expiry) >= Now Or .Expiry = "" Then
            addString .LocalSymbol, "Local Symbol"
        Else
            addString "", "Local Symbol"
        End If
    Else
        addString .LocalSymbol, "Local Symbol"
    End If
    If mServerVersion >= 31 Then
        If .Sectype = SecTypeCash Or .Sectype = SecTypeIndex Or .Sectype = SecTypeStock Then
            addBoolean False, "Include expired" ' can't include expired for non-expiring contracts
        Else
            addBoolean .IncludeExpired, "Include expired"
        End If
    End If
    
    
End With
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reqHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As HistoricalDataRequest)
Const ProcName As String = "reqHistoricalData"

On Error GoTo Err

Dim lTwsId As Long

Const Version = 4

If mConnectionState <> ConnectionStates.ConnConnected Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not connected"

lTwsId = getTwsHistRequestIdFromCallersRequestId(pRequestId)

logMessage "Requesting historical data for: " & pRequest.Contract.LocalSymbol & _
            "; id=" & lTwsId & _
            "; barsize=" & pRequest.BarSizeSetting & _
            "; endTime=" & pRequest.EndDateTime & _
            "; duration=" & pRequest.Duration, _
            ProcName

addMsgId REQ_HISTORICAL_DATA
addString Version, "Version"
addString lTwsId, "Request id"

With pRequest.Contract
    addString .Symbol, "Symbol"
    addString gTwsSecTypeToShortString(.Sectype), "Sec type"
    addString .Expiry, "Expiry"
    addDouble .Strike, "Strike"
    addString gTwsOptionRightToString(.OptRight), "Right"
    addString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    addString .Exchange, "Exchange"
    addString .PrimaryExch, "Primary Exchange"
    addString .CurrencyCode, "Currency"
    
    If .Sectype = TwsSecTypeCash Or _
        .Sectype = TwsSecTypeIndex Or _
        .Sectype = TwsSecTypeStock _
    Then
        addString .LocalSymbol, "Local Symbol"
        If mServerVersion >= 31 Then addString 0, "Include expired" ' can't include expired for non-expiring contracts
    Else
        If expiryToDate(.Expiry) < Now Then
            addString "", "Local Symbol"
        Else
            addString .LocalSymbol, "Local Symbol"
        End If
        If mServerVersion >= 31 Then addString 1, "Include expired"
    End If
End With

If mServerVersion >= 20 Then
    addString pRequest.EndDateTime, "End date"
    addString pRequest.BarSizeSetting, "Bar Size"
End If

addString pRequest.Duration, "Duration"
addString 0, "Use RTH"
addString pRequest.WhatToShow, "What to show"

If mServerVersion > 16 Then addString TwsHistDataDateFormats.DateFormatString, "Date format"

If pRequest.Contract.Sectype = TwsSecTypeCombo Then
    addString pRequest.Contract.ComboLegs.Count, "Combo legs count"
    Dim lComboLeg As TwsComboLeg
    Dim i As Long
    For Each lComboLeg In pRequest.Contract.ComboLegs
        With lComboLeg
            i = i + 1
            addString .ConId, "ConId" & i
            addString .Ratio, "Ratio" & i
            addString gTwsOrderActionToString(.Action), "Action" & i
            addString .Exchange, "Exchange" & i
        End With
    Next
End If

send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "reqMarketData"

Const Version = 9

On Error GoTo Err

logMessage "Requesting market data for: " & pContract.ToString, ProcName

If mServerVersion < MIN_SERVER_VER_SNAPSHOT_MKT_DATA And pSnapshot Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Snapshot market data requests not supported"

If mServerVersion < MIN_SERVER_VER_UNDER_COMP And pContract.UnderComp.ConId <> 0 Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Delta-neutral orders not supported"

If mServerVersion < MIN_SERVER_VER_REQ_MKT_DATA_CONID And pContract.ConId <> 0 Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "ConId property not supported"

addMsgId REQ_MKT_DATA
addString Version, "Version"
addString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
With pContract
    If mServerVersion >= MIN_SERVER_VER_REQ_MKT_DATA_CONID Then addString .ConId, "Con id"
    addString .Symbol, "Symbol"
    addString gTwsSecTypeToShortString(.Sectype), "Sec type"
    addString .Expiry, "Expiry"
    addDouble .Strike, "Strike"
    addString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then addString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    addString .Exchange, "Exchange"
    If mServerVersion >= 15 Then addString .PrimaryExch, "Primary Exchange"
    addString .CurrencyCode, "Currency"
    If mServerVersion >= 2 Then addString .LocalSymbol, "Local Symbol"
    
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .Sectype = SecurityTypes.SecTypeCombo Then
        addString .ComboLegs.Count, "Combolegs count"
        Dim comboLeg As TwsComboLeg
        Dim i As Long
        For Each comboLeg In .ComboLegs
            With comboLeg
                i = i + 1
                addString .ConId, "ConId" & i
                addString .Ratio, "Ratio" & i
                addString gTwsOrderActionToString(.Action), "Action" & i
                addString .Exchange, "Exchange" & i
            End With
        Next
    End If
    
    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
       If .UnderComp.ConId <> 0 Then
           addBoolean True, "Under comp"
           addString .UnderComp.ConId, "Under comp conid"
           addString .UnderComp.Delta, "Under comp delta"
           addString .UnderComp.Price, "Under comp price"
       Else
           addBoolean False, "Under comp"
       End If
    End If
    
    If mServerVersion >= 31 Then
         ' Note: Even though SHORTABLE tick type supported only
         '       starting server version 33 it would be relatively
         '       expensive to expose this restriction here.
         '
         '       Therefore we are relying on TWS doing validation.
         '
        addString pGenericTicks, "Generic tick list"
    End If
    
    If mServerVersion >= MIN_SERVER_VER_SNAPSHOT_MKT_DATA Then addBoolean pSnapshot, "Snapshot"
    
    send
End With

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub reqMarketDepth(ByVal TickerId As Long, _
                        ByVal pContract As TwsContract, _
                        Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "reqMarketDepth"


Const Version = 3


On Error GoTo Err

addMsgId REQ_MKT_DEPTH
addString Version, "Version"
addString getTwsMarketDepthRequestIdFromCallersRequestId(TickerId), "Request id"
With pContract
    addString .Symbol, "Symbol"
    addString gTwsSecTypeToShortString(.Sectype), "Sec type"
    addString .Expiry, "Expiry"
    addDouble .Strike, "Strike"
    addString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then addString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    addString .Exchange, "Exchange"
    addString .CurrencyCode, "Currency"
    addString .LocalSymbol, "Local Symbol"
    If mServerVersion >= 19 Then addString pNumberOfRows, "Num rows"
End With
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub retryConnection()
Const ProcName As String = "retryConnection"

On Error GoTo Err

If mConnectionRetryIntervalSecs <> 0 Then
    Set mConnectionTimer = CreateIntervalTimer(mConnectionRetryIntervalSecs, ExpiryTimeUnitSeconds)
    mConnectionTimer.StartTimer
    mRetryingConnection = True
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub rollbackBuffer()
'Debug.Print "Rollback: mInputIndex=" & mInputParseIndexChk
mInputParseIndex = mInputParseIndexChk
End Sub

Private Function SecTypeFromString(ByVal Value As String) As SecurityTypes
Select Case UCase$(Value)
Case "STK"
    SecTypeFromString = SecTypeStock
Case "FUT"
    SecTypeFromString = SecTypeFuture
Case "OPT"
    SecTypeFromString = SecTypeOption
Case "FOP"
    SecTypeFromString = SecTypeFuturesOption
Case "CASH"
    SecTypeFromString = SecTypeCash
Case "IND"
    SecTypeFromString = SecTypeIndex
End Select
End Function

Private Function SecTypeToString(ByVal Value As SecurityTypes) As String
Select Case Value
Case SecTypeStock
    SecTypeToString = "STK"
Case SecTypeFuture
    SecTypeToString = "FUT"
Case SecTypeOption
    SecTypeToString = "OPT"
Case SecTypeFuturesOption
    SecTypeToString = "FOP"
Case SecTypeCash
    SecTypeToString = "CASH"
Case SecTypeIndex
    SecTypeToString = "IND"
End Select
End Function

Private Sub send(Optional ByVal ignoreLogLevel As Boolean)
Const ProcName As String = "send"

On Error GoTo Err

logSocketOutputMessage ignoreLogLevel

ReDim Preserve mOutputBuffer(mOutputBufferIndex - 1) As Byte
If mSocketLogger.IsLoggable(LogLevelMediumDetail) Then
    mSocketLogger.Log LogLevels.LogLevelMediumDetail, _
                        showBuffer("Out buf: ", mOutputBuffer)
End If
mSocket.SendData mOutputBuffer
ReDim mOutputBuffer(15) As Byte
mOutputBufferIndex = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


Private Sub setTwsLogLevel(logLevel As TwsLogLevels)
Const ProcName As String = "setTwsLogLevel"


Const Version = 1


On Error GoTo Err

addMsgId SET_SERVER_LOGLEVEL
addString Version, "Version"
addString logLevel, "Loglevel"
send

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function showBuffer( _
                ByRef header As String, _
                ByRef buffer() As Byte) As String
Const ProcName As String = "showBuffer"

On Error GoTo Err

showBuffer = header & Replace(StrConv(buffer, vbUnicode), Chr$(0), "_")

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function stopTriggerMethodToString(ByVal Value As TwsStopTriggerMethods) As String
Select Case Value
Case TwsStopTriggerMethods.TwsStopTriggerBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerDefault, _
        TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerDoubleLast, _
        TwsStopTriggerMethods.TwsStopTriggerLast, _
        TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk, _
        TwsStopTriggerMethods.TwsStopTriggerMidPoint
    stopTriggerMethodToString = CStr(Value)
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid stop trigger method"
End Select
End Function

