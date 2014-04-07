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

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                    As String = "TwsAPI"

Private Const ClientVersion As Long = 47

'================================================================================
' Enums
'================================================================================

Private Enum TwsHistDataDateFormats
    DateFormatString = 1
    DateFormatInteger = 2
End Enum
'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mServer                                     As String
Private mPort                                       As String

Private mClientID                                   As Long

Private mProgramErrorHandler                        As ProgramErrorListener
Private mErrorAndNotificationConsumer               As IErrorAndNotificationConsumer
Private mConnectionStatusConsumer                   As IConnectionStatusConsumer

Private WithEvents mSocketHandler                   As SocketHandler
Attribute mSocketHandler.VB_VarHelpID = -1

Private mWriter                                     As BufferedWriter

Private WithEvents mInMessageHandler                As InputMessageHandler
Attribute mInMessageHandler.VB_VarHelpID = -1

Private mServerVersion                              As Long

Private mConnectionState                            As TwsConnectionStates

Private WithEvents mPerformanceTimerSecond          As IntervalTimer
Attribute mPerformanceTimerSecond.VB_VarHelpID = -1
Private WithEvents mPerformanceTimerPeriod          As IntervalTimer
Attribute mPerformanceTimerPeriod.VB_VarHelpID = -1

Private mPerformanceLogger                          As Logger

Private mStatsRecorder                              As PerformanceStatsRecorder

Private mConnectionRetryIntervalSecs                As Long

Private mLogLevel                                   As TwsLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

createLoggers

Set mInMessageHandler = New InputMessageHandler
Set mStatsRecorder = New PerformanceStatsRecorder

mConnectionState = TwsConnectionStates.TwsConnNotConnected

mLogLevel = 0

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Debug.Print "TwsAPI terminated"
End Sub

'================================================================================
' mInMessageHandler Event Handlers
'================================================================================

Private Sub mInMessageHandler_ApiConnectionEstablished(ByVal pServerVersion As Variant)
Const ProcName As String = "mInMessageHandler_ApiConnectionEstablished"
On Error GoTo Err

mServerVersion = pServerVersion
If mLogLevel <> 0 Then sendLogLevel

setConnectionState TwsConnConnected, ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mPerformanceTimerSecond Event Handlers
'================================================================================

Private Sub mPerformanceTimerPeriod_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mPerformanceTimerPeriod_TimerExpired"
On Error GoTo Err

mPerformanceLogger.Log LogLevelDetail, "Socket message statistics:" & vbCrLf & mStatsRecorder.GenerateStats

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mPerformanceTimerSecond_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mPerformanceTimerSecond_TimerExpired"
On Error GoTo Err

mStatsRecorder.AccumulateStats

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' mSocketHandler Event Handlers
'================================================================================

Private Sub mSocketHandler_Connected()
Const ProcName As String = "mSocketHandler_Connected"
On Error GoTo Err

gLog "Connecting to Tws: negotiating versions and client id", ModuleName, ProcName
mWriter.AddString ClientVersion, "Client version"
mWriter.Send
mWriter.AddString mClientID, "Client id"
mWriter.Send

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectFailed( _
                ByVal pDescription As String, _
                ByVal pRetryInterval As Long)
Const ProcName As String = "mSocketHandler_ConnectFailed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnFailed, pDescription & ": " & ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Connecting()
Const ProcName As String = "mSocketHandler_Connecting"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnConnecting, ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_ConnectionClosed( _
                ByVal pDescription As String)
Const ProcName As String = "mSocketHandler_ConnectionClosed"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pDescription & ": " & ConnectionString

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

Private Sub mSocketHandler_Disconnected(ByVal pReason As String)
Const ProcName As String = "mSocketHandler_Disconnected"
On Error GoTo Err

setConnectionState TwsConnectionStates.TwsConnNotConnected, pReason

Exit Sub

Err:
gNotifyUnhandledError mProgramErrorHandler, ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let AccountDataConsumer(ByVal Value As IAccountDataConsumer)
mInMessageHandler.AccountDataConsumer = Value
End Property

Public Property Get AccountDataConsumer() As IAccountDataConsumer
Set AccountDataConsumer = mInMessageHandler.AccountDataConsumer
End Property

Public Property Get ClientID() As Long
ClientID = mClientID
End Property

Public Property Let ConnectionRetryIntervalSecs(ByVal Value As Long)
Const ProcName As String = "ConnectionRetryIntervalSecs"
On Error GoTo Err

AssertArgument Value >= 0, "Value cannot be negative"
mConnectionRetryIntervalSecs = Value
If Not mSocketHandler Is Nothing Then mSocketHandler.ConnectionRetryIntervalSecs = Value

Exit Property

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Property

Public Property Get ConnectionRetryIntervalSecs() As Long
ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Get ConnectionState() As TwsConnectionStates
ConnectionState = mConnectionState
End Property

Public Property Let ConnectionStatusConsumer(ByVal Value As IConnectionStatusConsumer)
Set mConnectionStatusConsumer = Value
mInMessageHandler.ConnectionStatusConsumer = Value
End Property

Public Property Get ConnectionStatusConsumer() As IConnectionStatusConsumer
Set ConnectionStatusConsumer = mConnectionStatusConsumer
End Property

Public Property Get ConnectionString() As String
ConnectionString = "server=" & mServer & _
                        " port=" & mPort & _
                        " client Id=" & mClientID
End Property

Public Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
mInMessageHandler.ContractDetailsConsumer = Value
End Property

Public Property Get ContractDetailsConsumer() As IContractDetailsConsumer
Set ContractDetailsConsumer = mInMessageHandler.ContractDetailsConsumer
End Property

Public Property Let ErrorAndNotificationConsumer(ByVal Value As IErrorAndNotificationConsumer)
Set mErrorAndNotificationConsumer = Value
End Property

Public Property Get ErrorAndNotificationConsumer() As IErrorAndNotificationConsumer
Set ErrorAndNotificationConsumer = mErrorAndNotificationConsumer
End Property

Public Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
mInMessageHandler.HistDataConsumer = Value
End Property

Public Property Get HistDataConsumer() As IHistDataConsumer
Set HistDataConsumer = mInMessageHandler.HistDataConsumer
End Property

Public Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
mInMessageHandler.MarketDataConsumer = Value
End Property

Public Property Get MarketDataConsumer() As IMarketDataConsumer
Set MarketDataConsumer = mInMessageHandler.MarketDataConsumer
End Property

Public Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
mInMessageHandler.MarketDepthConsumer = Value
End Property

Public Property Get MarketDepthConsumer() As IMarketDepthConsumer
Set MarketDepthConsumer = mInMessageHandler.MarketDepthConsumer
End Property

Public Property Get MaxContractRequestId() As Long
MaxContractRequestId = BaseOrderId - BaseContractRequestId - 1
End Property

Public Property Get MaxExecutionsRequestId() As Long
MaxExecutionsRequestId = BaseContractRequestId - BaseExecutionsRequestId - 1
End Property

Public Property Get MaxHistoricalDataRequestId() As Long
MaxHistoricalDataRequestId = BaseExecutionsRequestId - BaseHistoricalDataRequestId - 1
End Property

Public Property Get MaxMarketDataRequestId() As Long
MaxMarketDataRequestId = BaseMarketDepthRequestId - 1
End Property

Public Property Get MaxMarketDepthRequestId() As Long
MaxMarketDepthRequestId = BaseHistoricalDataRequestId - BaseMarketDepthRequestId - 1
End Property

Public Property Get MaxOrderId() As Long
MaxOrderId = &H7FFFFFFF
End Property

Public Property Get NextOrderId() As Long
NextOrderId = mInMessageHandler.NextOrderId
End Property

Public Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
mInMessageHandler.OrderInfoConsumer = Value
End Property

Public Property Get OrderInfoConsumer() As IOrderInfoConsumer
Set OrderInfoConsumer = mInMessageHandler.OrderInfoConsumer
End Property

Public Property Get Port() As String
Port = mPort
End Property

Public Property Let ProgramErrorHandler(ByVal Value As ProgramErrorListener)
Set mProgramErrorHandler = Value
End Property

Public Property Get ProgramErrorHandler() As ProgramErrorListener
Set ProgramErrorHandler = mProgramErrorHandler
End Property

Public Property Let ScannerDataConsumer(ByVal Value As IScannerDataConsumer)
mInMessageHandler.ScannerDataConsumer = Value
End Property

Public Property Get ScannerDataConsumer() As IScannerDataConsumer
Set ScannerDataConsumer = mInMessageHandler.ScannerDataConsumer
End Property

Public Property Get SerialiseContractDetailsRequests() As Boolean
SerialiseContractDetailsRequests = mInMessageHandler.SerialiseContractDetailsRequests
End Property

Public Property Get Server() As String
Server = mServer
End Property

Public Property Get ServerVersion() As Long
ServerVersion = mServerVersion
End Property

'================================================================================
' Methods
'================================================================================

Public Sub CancelHistoricalData( _
                ByVal pRequestId As Long)
Const ProcName As String = "CancelHistoricalData"
On Error GoTo Err

cancelHistDataRequest pRequestId

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketData(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

cancelMarketDataEx pTickerId

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelMarketDepth(ByVal pTickerId As Long)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const Version = 1

mWriter.AddMessageId CANCEL_MKT_DEPTH
mWriter.AddString Version, "Version"
mWriter.AddString getTwsMarketDepthRequestIdFromCallersRequestId(pTickerId), "Ticker id"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mErrorAndNotificationConsumer, pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub CancelNewsBulletins()
Const ProcName As String = "CancelNewsBulletins"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const Version = 1

mWriter.AddMessageId CANCEL_NEWS_BULLETINS
mWriter.AddString Version, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub CancelOrder(ByVal pOrderId As Long)
Const ProcName As String = "CancelOrder"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1
mWriter.AddMessageId CANCEL_ORDER
mWriter.AddString Version, "Version"
mWriter.AddString pOrderId, "Order id"
mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Connect()
Const ProcName As String = "Connect"
On Error GoTo Err

Dim lReader As BufferedReader

Set mSocketHandler = New SocketHandler
Set lReader = New BufferedReader
mSocketHandler.Initialise mServer, mPort, lReader, mProgramErrorHandler
mSocketHandler.ConnectionRetryIntervalSecs = mConnectionRetryIntervalSecs

Set mWriter = New BufferedWriter
mWriter.Initialise mSocketHandler

mInMessageHandler.Initialise lReader, mWriter, mStatsRecorder, mErrorAndNotificationConsumer, mProgramErrorHandler

mSocketHandler.Connect

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub Disconnect( _
                ByVal pReason As String)
Const ProcName As String = "Disconnect"
On Error GoTo Err

mSocketHandler.Disconnect pReason
Set mSocketHandler = Nothing

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Function GetRawTwsMessage() As Byte()
Const ProcName As String = "GetRawMessage"
On Error GoTo Err

GetRawTwsMessage = mInMessageHandler.BufferedReader.GetRawMessage

Exit Function

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long, _
                ByVal pClientId As Long)
If pServer = "" Then pServer = "127.0.0.1"
mServer = pServer
mPort = pPort
mClientID = pClientId
End Sub

Public Sub PlaceOrder( _
                ByVal pOrder As TwsOrder, _
                ByVal pContract As TwsContract)
Const ProcName As String = "PlaceOrder"
On Error GoTo Err

Dim i As Long

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 30
mWriter.AddMessageId PLACE_ORDER
mWriter.AddString Version, "Version"

If pOrder.OrderId = 0 Then pOrder.OrderId = NextOrderId
Assert pOrder.OrderId >= BaseOrderId, "Order id must not be less than " & BaseOrderId
If pOrder.OrderId > mInMessageHandler.NextOrderId Then mInMessageHandler.NextOrderId = pOrder.OrderId + 1

mWriter.AddString pOrder.OrderId, "Order id"

' mwriter.send contract fields
With pContract
    If mServerVersion >= MIN_SERVER_VER_PLACE_ORDER_CONID Then mWriter.AddString .ConId, "Con id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sectype"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier" ' multiplier - this is the number of units of the underlying
                                                            ' delivered per contract - see http://www.interactivebrokers.com/discus/messages/2/27413.html
    mWriter.AddString .Exchange, "Exchange"
    If mServerVersion >= 14 Then mWriter.AddString .PrimaryExch, "Primary Exchange"    ' only relevant when
                                                        ' Exchange is SMART and there are SMART routers in more
                                                        ' than one country (eg try IBM in Tws)
    mWriter.AddString .CurrencyCode, "Currency"
    If mServerVersion >= 2 Then mWriter.AddString .LocalSymbol, "Local Symbol"
    If mServerVersion >= MIN_SERVER_VER_SEC_ID_TYPE Then
          mWriter.AddString .SecIdType, "Sec id type"
          mWriter.AddString .SecId, "Sec id"
    End If
End With

With pOrder
    
    ' mwriter.send main Order fields
    mWriter.AddString gTwsOrderActionToString(.Action), "Action"
    mWriter.AddString .TotalQuantity, "Quantity"
    mWriter.AddString gTwsOrderTypeToString(.OrderType), "Order type"
    If mServerVersion <= 10 And _
        (UCase$(pContract.Exchange) = "LIFFE" Or _
        UCase$(pContract.Exchange) = "LSE") _
    Then
        mWriter.AddDouble .LmtPrice / 100, "Price"
        mWriter.AddDouble .AuxPrice / 100, "Aux price"
    Else
        mWriter.AddDouble .LmtPrice, "Price"
        mWriter.AddDouble .AuxPrice, "Aux price"
    End If
    
    ' mwriter.send extended Order fields
    mWriter.AddString gTwsOrderTIFToString(.Tif), "TIF"
    mWriter.AddString .OcaGroup, "Oca Group"
    mWriter.AddString .Account, "Account"
    mWriter.AddString .OpenClose, "OpenClose"
    mWriter.AddString .Origin, "Origin"
    mWriter.AddString .OrderRef, "Order ref"
    mWriter.AddBoolean .Transmit, "Transmit"
    If mServerVersion >= 4 Then mWriter.AddString .ParentId, "Parent id"
    If mServerVersion >= 5 Then
        mWriter.AddBoolean .BlockOrder, "Block Order"
        mWriter.AddBoolean .SweepToFill, "Sweep to fill"
        mWriter.AddString .DisplaySize, "Display Size"
        mWriter.AddString stopTriggerMethodToString(.TriggerMethod), "Trigger method"
        If mServerVersion < 38 Then
            mWriter.AddBoolean False, "Outside RTH"
        Else
            mWriter.AddBoolean .OutsideRth, "Outside RTH"
        End If
    End If
    If mServerVersion >= 7 Then mWriter.AddBoolean .Hidden, "Hidden"
End With

' mwriter.send combo legs for BAG requests
With pContract
    If mServerVersion >= 8 And .Sectype = TwsSecTypeCombo Then
        If .ComboLegs.Count = 0 Then
            mWriter.AddString 0, "Combo legs count"
        Else
            mWriter.AddString .ComboLegs.Count, "Combo legs count"
            Dim comboLeg As TwsComboLeg
            For Each comboLeg In .ComboLegs
                With comboLeg
                    mWriter.AddString .ConId, "Leg " & i & " Con id"
                    mWriter.AddString .Ratio, "Leg " & i & " Ration"
                    mWriter.AddString gTwsOrderActionToString(.Action), "Leg " & i & " Action"
                    mWriter.AddString .Exchange, "Leg " & i & " Exchange"
                    mWriter.AddString legOpenCloseToString(.OpenClose), "Leg " & i & " Open/close"
                End With
            Next
        End If
    End If
End With

With pOrder
    If mServerVersion >= 9 Then mWriter.AddString "", "Shares Allocation"
    
    If mServerVersion >= 10 Then mWriter.AddString .DiscretionaryAmt, "Discretionary amount"
    
    If mServerVersion >= 11 Then mWriter.AddString .GoodAfterTime, "Good after Time"
    
    If mServerVersion >= 12 Then mWriter.AddString .GoodTillDate, "Good till date"
    
    If mServerVersion >= 13 Then
        mWriter.AddString .FaGroup, "FAGroup"
        mWriter.AddString .FaMethod, "FAMethod"
        mWriter.AddString .FaPercentage, "FAPercentage"
        mWriter.AddString .FaProfile, "FAProfile"
    End If
    
    If mServerVersion >= 18 Then 'institutional short sale slot fields.
        mWriter.AddString .ShortSaleSlot, " Short Sale Slot"           ' 0 only for retail, 1 or 2 only for institution.
        mWriter.AddString .DesignatedLocation, "Designated Location"       ' only populate when shortSaleSlot = 2.
    End If
    
    If mServerVersion >= 19 Then
        mWriter.AddString .OcaType, "Oca type"
        If mServerVersion < 38 Then
            mWriter.AddBoolean False, "RTH only" ' (now combined with ignoreRTH)
        End If
        mWriter.AddString .Rule80A, "Rule 80A"
        mWriter.AddString .SettlingFirm, "Settling firm"
        mWriter.AddBoolean .AllOrNone, "All or none"
        mWriter.AddLongMax .MinQty, "Minimum quantity"
        mWriter.AddDoubleMax .PercentOffset, "Percent Offset"
        mWriter.AddBoolean .ETradeOnly, "E-trade only"
        mWriter.AddBoolean .FirmQuoteOnly, "Firm quote only"
        mWriter.AddDoubleMax .NbboPriceCap, "NBBO price cap"
        mWriter.AddLongMax .AuctionStrategy, "Auction strategy"
        mWriter.AddDoubleMax .StartingPrice, "Starting price"
        mWriter.AddDoubleMax .StockRefPrice, "Stock ref price"
        mWriter.AddDoubleMax .Delta, "Delta"
        
        If mServerVersion = 26 And .OrderType = TwsOrderTypeVol Then
            ' Volatility orders had specific watermark price attribs in server version 26
            mWriter.AddDoubleMax MaxDouble, "Stock range lower"
            mWriter.AddDoubleMax MaxDouble, "Stock range upper"
        Else
            mWriter.AddDoubleMax .StockRangeLower, "Stock range lower"
            mWriter.AddDoubleMax .StockRangeUpper, "Stock range upper"
        End If
    End If
    
    If mServerVersion >= 22 Then mWriter.AddBoolean .OverridePercentageConstraints, "Override percentage constraints"

    If mServerVersion >= 26 Then ' Volatility orders
        mWriter.AddDoubleMax .Volatility, "Volatility"
        mWriter.AddLongMax .VolatilityType, "Volatility type"
        If mServerVersion < 28 Then
            mWriter.AddBoolean (.DeltaNeutralOrderType = TwsOrderTypeMarket), "Delta neutral Order type"
        Else
            mWriter.AddString gTwsOrderTypeToString(.DeltaNeutralOrderType), "Delta neutral Order type"
            mWriter.AddDoubleMax .DeltaNeutralAuxPrice, "Delta neutral aux price"
        End If
        
        mWriter.AddBoolean .ContinuousUpdate, "Continuous update"
        
        If mServerVersion = 26 Then
            If .OrderType = TwsOrderTypeVol Then
                ' Volatility orders had specific watermark price attribs in server version 26
                mWriter.AddDoubleMax .StockRangeLower, "Stock range lower"
                mWriter.AddDoubleMax .StockRangeUpper, "Stock range upper"
            Else
                mWriter.AddDoubleMax MaxDouble, "Stock range lower"
                mWriter.AddDoubleMax MaxDouble, "Stock range upper"
            End If
        End If
        
        mWriter.AddLongMax .ReferencePriceType, "Reference price type"
    End If
    
    If mServerVersion >= 30 Then mWriter.AddDoubleMax .TrailStopPrice, "Trail stop price"

    If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS) Then
        If (mServerVersion >= MIN_SERVER_VER_SCALE_ORDERS2) Then
            mWriter.AddLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
            mWriter.AddLongMax .ScaleSubsLevelSize, "ScaleSubsLevelSize"
        
        Else
            mWriter.AddString "", "ScaleSubsLevelSize"
            mWriter.AddLongMax .ScaleInitLevelSize, "ScaleInitLevelSize"
            
        End If
        mWriter.AddDoubleMax .ScalePriceIncrement, "ScalePriceIncrement"
    End If
    
    If (mServerVersion >= MIN_SERVER_VER_PTA_ORDERS) Then
        mWriter.AddString .ClearingAccount, "Clearing account"
        mWriter.AddString .ClearingIntent, "Clearing intent"
    End If

    If (mServerVersion >= MIN_SERVER_VER_NOT_HELD) Then
        mWriter.AddBoolean .NotHeld, "Not held"
    End If

    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
        If pContract.UnderComp.ConId <> 0 Then
            mWriter.AddBoolean True, "Under comp"
            mWriter.AddString pContract.UnderComp.ConId, "Under comp conid"
            mWriter.AddDouble pContract.UnderComp.Delta, "Under comp delta"
            mWriter.AddDouble pContract.UnderComp.Price, "Under comp price"
        Else
            mWriter.AddBoolean False, "Under comp"
        End If
    End If

    If (mServerVersion >= MIN_SERVER_VER_ALGO_ORDERS) Then
        mWriter.AddString .AlgoStrategy, "Algo strategy"
        If .AlgoStrategy <> "" Then
            Dim algoParamsCount  As Long
            
            On Error Resume Next
            algoParamsCount = UBound(.AlgoParams) + 1
            On Error GoTo Err
             
            mWriter.AddString algoParamsCount, "Algo params count"
            If algoParamsCount > 0 Then
                Dim lAlgoParams() As TwsTagValue
                lAlgoParams = pOrder.AlgoParams
                For i = 0 To algoParamsCount - 1
                    mWriter.AddString lAlgoParams(i).Tag, "Tag" & i
                    mWriter.AddString lAlgoParams(i).Value, "value" & i
                Next
            End If
        End If
    End If

    If (mServerVersion >= MIN_SERVER_VER_WHAT_IF_ORDERS) Then
        mWriter.AddBoolean .WhatIf, "WhatIf"
    End If
End With

mWriter.Send True

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub
                    
Public Sub ProcessRawTwsMessage(ByRef pMessage() As Byte, ByVal pLog As Boolean)
Const ProcName As String = "ProcessRawTwsMessage"
On Error GoTo Err

Dim lReader As New BufferedReader
Dim lInMessageHandler As New InputMessageHandler

lReader.AddData pMessage
lInMessageHandler.Initialise lReader, mWriter, mStatsRecorder, mErrorAndNotificationConsumer, mProgramErrorHandler
lInMessageHandler.AccountDataConsumer = mInMessageHandler.AccountDataConsumer
lInMessageHandler.ConnectionStatusConsumer = mInMessageHandler.ConnectionStatusConsumer
lInMessageHandler.ContractDetailsConsumer = mInMessageHandler.ContractDetailsConsumer
lInMessageHandler.HistDataConsumer = mInMessageHandler.HistDataConsumer
lInMessageHandler.MarketDataConsumer = mInMessageHandler.MarketDataConsumer
lInMessageHandler.MarketDepthConsumer = mInMessageHandler.MarketDepthConsumer
lInMessageHandler.OrderInfoConsumer = mInMessageHandler.OrderInfoConsumer
lInMessageHandler.ScannerDataConsumer = mInMessageHandler.ScannerDataConsumer
lInMessageHandler.ProcessBuffer pLog

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Public Sub ReplaceFA(ByVal DataType As TwsFADataTypes, ByVal xml As String)
Const ProcName As String = "ReplaceFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1

Assert mServerVersion >= 13, "Server version does not support ReplaceFA", ErrorCodes.ErrUnsupportedOperationException

mWriter.AddMessageId REPLACE_FA
mWriter.AddString Version, "Version"
mWriter.AddString DataType, "Data type"
mWriter.AddString xml, "XML"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAccountUpdates(ByVal subscribe As Boolean, ByVal acctCode As String)
Const ProcName As String = "RequestAccountUpdates"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 2

mWriter.AddMessageId REQ_ACCT_DATA
mWriter.AddString Version, "Version"
mWriter.AddBoolean subscribe, "Subscribe"

If mServerVersion >= 9 Then mWriter.AddString acctCode, "Account code"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAllOpenOrders()
Const ProcName As String = "RequestAllOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
                            

Const Version = 1

mWriter.AddMessageId REQ_ALL_OPEN_ORDERS
mWriter.AddString Version, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestAutoOpenOrders(ByVal autoBind As Boolean)
Const ProcName As String = "RequestAutoOpenOrders"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1

mWriter.AddMessageId REQ_AUTO_OPEN_ORDERS
mWriter.AddString Version, "Version"
mWriter.AddBoolean autoBind, "Autobind"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "RequestContractDetails"
On Error GoTo Err

reqContractDetails pRequestId, pContract

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestExecutions( _
                ByVal pRequestId As Long, _
                ByVal filter As TwsExecutionFilter)
Dim theFilter As TwsExecutionFilter

Const ProcName As String = "RequestExecutions"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 3

mWriter.AddMessageId REQ_EXECUTIONS
mWriter.AddString Version, "Version"

If mServerVersion >= MIN_SERVER_VER_EXECUTION_DATA_CHAIN Then mWriter.AddString getTwsExecutionsRequestIdFromCallersRequestId(pRequestId), "ReqId"

If mServerVersion >= 9 Then
    If filter Is Nothing Then
        Set theFilter = New TwsExecutionFilter
    Else
        Set theFilter = filter
    End If
    
    With theFilter
        mWriter.AddString .ClientID, "Client id"
        mWriter.AddString .AccountCode, "Account code"
        If .Time = Int(.Time) Then
            mWriter.AddString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd"), ""), "Fill Time"
        Else
            mWriter.AddString IIf(CLng(.Time) <> 0, Format(.Time, "yyyymmdd-hh\:nn\:ss"), ""), "Fill Time"
        End If
        mWriter.AddString .Symbol, "Symbol"
        mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
        mWriter.AddString .Exchange, "Exchange"
        mWriter.AddString gTwsOrderActionToString(.Action), "Action"
    End With
End If

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestFA(ByVal DataType As TwsFADataTypes)
Const ProcName As String = "RequestFA"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1

Assert mServerVersion >= 13, "Server version does not support RequestFA", ErrorCodes.ErrUnsupportedOperationException

mWriter.AddMessageId REQ_FA
mWriter.AddString Version, "Version"
mWriter.AddString DataType, "Data type"

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestManagedAccounts()
Const ProcName As String = "RequestManagedAccounts"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1

mWriter.AddMessageId REQ_MANAGED_ACCTS
mWriter.AddString Version, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

reqMarketData pTickerId, pContract, pGenericTicks, pSnapshot

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestMarketDepth( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

reqMarketDepth pTickerId, pContract, pNumberOfRows

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest)
Const ProcName As String = "RequestHistoricalData"
On Error GoTo Err

reqHistoricalData pRequestId, pRequest

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestNewsBulletins(allMsgs As Boolean)
Const ProcName As String = "RequestNewsBulletins"
On Error GoTo Err

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Const Version = 1

mWriter.AddMessageId REQ_NEWS_BULLETINS
mWriter.AddString Version, "Version"
mWriter.AddString allMsgs, "All messages"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub RequestOpenOrders()
Const ProcName As String = "RequestOpenOrders"
On Error GoTo Err

Const Version = 1

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

mWriter.AddMessageId REQ_OPEN_ORDERS
mWriter.AddString Version, "Version"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError mProgramErrorHandler, ProcName, ModuleName
End Sub

Public Sub SetTwsLogLevel(ByVal pLogLevel As TwsLogLevels)
mLogLevel = pLogLevel
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelHistDataRequest( _
                ByVal pRequestId As Long)
Const ProcName As String = "cancelHistDataRequest"
On Error GoTo Err

If mServerVersion < 24 Then
    ' cancel not supported by Tws
    Exit Sub
End If

If mConnectionState = TwsConnConnected Then
    Const Version As Long = 1

    mWriter.AddMessageId CANCEL_HISTORICAL_DATA
    mWriter.AddString Version, "Version"
    mWriter.AddString getTwsHistRequestIdFromCallersRequestId(pRequestId), "Request id"
    mWriter.Send
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub cancelMarketDataEx(ByVal pTickerId As Long)
Const ProcName As String = "cancelMarketDataEx"
On Error GoTo Err

If mConnectionState <> TwsConnectionStates.TwsConnConnected Then Exit Sub

Const Version = 2

mWriter.AddMessageId CANCEL_MKT_DATA
mWriter.AddString Version, "Version"
mWriter.AddString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub createLoggers()

Const ProcName As String = "createLoggers"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelDetail) Then
    If mPerformanceTimerSecond Is Nothing Then
        Set mPerformanceTimerSecond = CreateIntervalTimer(CDate(Int((Now + 2 * OneSecond) / OneSecond) * OneSecond), _
                                                ExpiryTimeUnitDateTime, _
                                                1000)
        Set mPerformanceTimerPeriod = CreateIntervalTimer(CDate(Int((Now + OneMinute) / OneMinute) * OneMinute), _
                                                ExpiryTimeUnitDateTime, _
                                                60000)
        mPerformanceTimerSecond.StartTimer
        mPerformanceTimerPeriod.StartTimer
        
        Set mPerformanceLogger = GetLogger("tradebuild.log.ibapi.performance")
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
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
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Function

Private Function getTwsContractRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersContractRequestId, "Max request id is " & MaxCallersContractRequestId
getTwsContractRequestIdFromCallersRequestId = pId + BaseContractRequestId
End Function

Private Function getTwsExecutionsRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersExecutionsRequestId, "Max request id is " & MaxCallersExecutionsRequestId
getTwsExecutionsRequestIdFromCallersRequestId = pId + BaseExecutionsRequestId
End Function

Private Function getTwsHistRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersHistoricalDataRequestId, "Max request id is " & MaxCallersHistoricalDataRequestId
getTwsHistRequestIdFromCallersRequestId = pId + BaseHistoricalDataRequestId
End Function

Private Function getTwsMarketDataRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDataRequestId, "Max request id is " & MaxCallersMarketDataRequestId
getTwsMarketDataRequestIdFromCallersRequestId = pId + BaseMarketDataRequestId
End Function

Private Function getTwsMarketDepthRequestIdFromCallersRequestId(ByVal pId As Long) As Long
AssertArgument pId <= MaxCallersMarketDepthRequestId, "Max request id is " & MaxCallersMarketDepthRequestId
getTwsMarketDepthRequestIdFromCallersRequestId = pId + BaseMarketDepthRequestId
End Function

Private Function legOpenCloseToString(ByVal Value As TwsLegOpenCloseCodes) As String
Select Case Value
Case TwsLegOpenCloseCodes.TwsLegOpenCloseSame, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseOpen, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseClose, _
        TwsLegOpenCloseCodes.TwsLegOpenCloseUnknown
    legOpenCloseToString = CStr(Value)
Case Else
    Assert False, "Invalid leg open/close value"
End Select
End Function

Private Sub reqContractDetails( _
                ByVal pRequestId As Long, _
                ByVal pContract As TwsContract)
Const ProcName As String = "reqContractDetails"
On Error GoTo Err

Const Version = 5

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"
AssertArgument pContract.Sectype <> TwsSecTypeCombo, "Combo contracts are not supported"

If SerialiseContractDetailsRequests Then
    Assert Not mInMessageHandler.ContractDetailsRequested, "Contract details requests must be serialised with this version of Tws"
    mInMessageHandler.ContractDetailsRequestStarted
End If

mWriter.AddMessageId REQ_CONTRACT_DATA
mWriter.AddString Version, "Version"

If mServerVersion >= MIN_SERVER_VER_CONTRACT_DATA_CHAIN Then
    mWriter.AddString getTwsContractRequestIdFromCallersRequestId(pRequestId), "Request id"
End If

With pContract
    If mServerVersion >= MIN_SERVER_VER_CONTRACT_CONID Then mWriter.AddString .ConId, "Contract id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString IIf(.Exchange = "", "*", .Exchange), "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    If (.Sectype = TwsSecTypeFuture Or _
        .Sectype = TwsSecTypeFuturesOption Or _
        .Sectype = TwsSecTypeOption) _
    Then
        If expiryToDate(.Expiry) >= Now Or .Expiry = "" Then
            mWriter.AddString .LocalSymbol, "Local Symbol"
        Else
            mWriter.AddString "", "Local Symbol"
        End If
    Else
        mWriter.AddString .LocalSymbol, "Local Symbol"
    End If
    If mServerVersion >= 31 Then
        If .Sectype = TwsSecTypeCash Or .Sectype = TwsSecTypeIndex Or .Sectype = TwsSecTypeStock Then
            mWriter.AddBoolean False, "Include expired" ' can't include expired for non-expiring contracts
        Else
            mWriter.AddBoolean .IncludeExpired, "Include expired"
        End If
    End If
    
    
End With
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub reqHistoricalData( _
                ByVal pRequestId As Long, _
                ByRef pRequest As TwsHistoricalDataRequest)
Const ProcName As String = "reqHistoricalData"
On Error GoTo Err

Const Version = 4

Assert mConnectionState = TwsConnectionStates.TwsConnConnected, "Not connected"

Dim lTwsId As Long
lTwsId = getTwsHistRequestIdFromCallersRequestId(pRequestId)

gLog "Requesting historical data for: " & pRequest.contract.LocalSymbol & _
            "; id=" & lTwsId & _
            "; barsize=" & pRequest.BarSizeSetting & _
            "; endTime=" & pRequest.EndDateTime & _
            "; duration=" & pRequest.Duration, _
            ModuleName, ProcName

mWriter.AddMessageId REQ_HISTORICAL_DATA
mWriter.AddString Version, "Version"
mWriter.AddString lTwsId, "Request id"

With pRequest.contract
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    
    If .Sectype = TwsSecTypeCash Or _
        .Sectype = TwsSecTypeIndex Or _
        .Sectype = TwsSecTypeStock _
    Then
        mWriter.AddString .LocalSymbol, "Local Symbol"
        If mServerVersion >= 31 Then mWriter.AddString 0, "Include expired" ' can't include expired for non-expiring contracts
    Else
        If expiryToDate(.Expiry) < Now Then
            mWriter.AddString "", "Local Symbol"
        Else
            mWriter.AddString .LocalSymbol, "Local Symbol"
        End If
        If mServerVersion >= 31 Then mWriter.AddString 1, "Include expired"
    End If
End With

If mServerVersion >= 20 Then
    mWriter.AddString pRequest.EndDateTime, "End date"
    mWriter.AddString pRequest.BarSizeSetting, "Bar Size"
End If

mWriter.AddString pRequest.Duration, "Duration"
mWriter.AddString 0, "Use RTH"
mWriter.AddString pRequest.WhatToShow, "What to show"

If mServerVersion > 16 Then mWriter.AddString TwsHistDataDateFormats.DateFormatString, "Date format"

If pRequest.contract.Sectype = TwsSecTypeCombo Then
    mWriter.AddString pRequest.contract.ComboLegs.Count, "Combo legs count"
    Dim lComboLeg As TwsComboLeg
    Dim i As Long
    For Each lComboLeg In pRequest.contract.ComboLegs
        With lComboLeg
            i = i + 1
            mWriter.AddString .ConId, "ConId" & i
            mWriter.AddString .Ratio, "Ratio" & i
            mWriter.AddString gTwsOrderActionToString(.Action), "Action" & i
            mWriter.AddString .Exchange, "Exchange" & i
        End With
    Next
End If

mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub reqMarketData( _
                ByVal pTickerId As Long, _
                ByVal pContract As TwsContract, _
                ByVal pGenericTicks As String, _
                ByVal pSnapshot As Boolean)
Const ProcName As String = "reqMarketData"

Const Version = 9

On Error GoTo Err

gLog "Requesting market data for: " & pContract.ToString, ModuleName, ProcName

Assert Not (mServerVersion < MIN_SERVER_VER_SNAPSHOT_MKT_DATA And pSnapshot), "Snapshot market data requests not supported"
Assert Not (mServerVersion < MIN_SERVER_VER_UNDER_COMP And pContract.UnderComp.ConId <> 0), "Delta-neutral orders not supported"
Assert Not (mServerVersion < MIN_SERVER_VER_REQ_MKT_DATA_CONID And pContract.ConId <> 0), "ConId property not supported"

mWriter.AddMessageId REQ_MKT_DATA
mWriter.AddString Version, "Version"
mWriter.AddString getTwsMarketDataRequestIdFromCallersRequestId(pTickerId), "Ticker id"
With pContract
    If mServerVersion >= MIN_SERVER_VER_REQ_MKT_DATA_CONID Then mWriter.AddString .ConId, "Con id"
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    If mServerVersion >= 15 Then mWriter.AddString .PrimaryExch, "Primary Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    If mServerVersion >= 2 Then mWriter.AddString .LocalSymbol, "Local Symbol"
    
    ' Add combo legs for BAG requests
    If mServerVersion >= 8 And .Sectype = TwsSecTypes.TwsSecTypeCombo Then
        mWriter.AddString .ComboLegs.Count, "Combolegs count"
        Dim comboLeg As TwsComboLeg
        Dim i As Long
        For Each comboLeg In .ComboLegs
            With comboLeg
                i = i + 1
                mWriter.AddString .ConId, "ConId" & i
                mWriter.AddString .Ratio, "Ratio" & i
                mWriter.AddString gTwsOrderActionToString(.Action), "Action" & i
                mWriter.AddString .Exchange, "Exchange" & i
            End With
        Next
    End If
    
    If (mServerVersion >= MIN_SERVER_VER_UNDER_COMP) Then
       If .UnderComp.ConId <> 0 Then
           mWriter.AddBoolean True, "Under comp"
           mWriter.AddString .UnderComp.ConId, "Under comp conid"
           mWriter.AddString .UnderComp.Delta, "Under comp delta"
           mWriter.AddString .UnderComp.Price, "Under comp price"
       Else
           mWriter.AddBoolean False, "Under comp"
       End If
    End If
    
    If mServerVersion >= 31 Then
         ' Note: Even though SHORTABLE tick type supported only
         '       starting server version 33 it would be relatively
         '       expensive to expose this restriction here.
         '
         '       Therefore we are relying on Tws doing validation.
         '
        mWriter.AddString pGenericTicks, "Generic tick list"
    End If
    
    If mServerVersion >= MIN_SERVER_VER_SNAPSHOT_MKT_DATA Then mWriter.AddBoolean pSnapshot, "Snapshot"
    
    mWriter.Send
End With

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName

End Sub

Private Sub reqMarketDepth(ByVal TickerId As Long, _
                        ByVal pContract As TwsContract, _
                        Optional ByVal pNumberOfRows As Long = 20)
Const ProcName As String = "reqMarketDepth"


Const Version = 3


On Error GoTo Err

mWriter.AddMessageId REQ_MKT_DEPTH
mWriter.AddString Version, "Version"
mWriter.AddString getTwsMarketDepthRequestIdFromCallersRequestId(TickerId), "Request id"
With pContract
    mWriter.AddString .Symbol, "Symbol"
    mWriter.AddString gTwsSecTypeToShortString(.Sectype), "Sec type"
    mWriter.AddString .Expiry, "Expiry"
    mWriter.AddDouble .Strike, "Strike"
    mWriter.AddString gTwsOptionRightToString(.OptRight), "Right"
    If mServerVersion >= 15 Then mWriter.AddString IIf(.Multiplier = 0, "", CStr(.Multiplier)), "Multiplier"
    mWriter.AddString .Exchange, "Exchange"
    mWriter.AddString .CurrencyCode, "Currency"
    mWriter.AddString .LocalSymbol, "Local Symbol"
    If mServerVersion >= 19 Then mWriter.AddString pNumberOfRows, "Num rows"
End With
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub sendLogLevel()
Const ProcName As String = "sendLogLevel"
On Error GoTo Err

Const Version = 1

mWriter.AddMessageId SET_SERVER_LOGLEVEL
mWriter.AddString Version, "Version"
mWriter.AddString mLogLevel, "Loglevel"
mWriter.Send

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

Private Sub setConnectionState(ByVal pState As TwsConnectionStates, ByVal pMessage As String)
Const ProcName As String = "setConnectionState"
On Error GoTo Err

mConnectionState = pState
If Not mConnectionStatusConsumer Is Nothing Then mConnectionStatusConsumer.NotifyAPIConnectionStateChange mConnectionState, pMessage

Exit Sub

Err:
gHandleUnexpectedError Nothing, ProcName, ModuleName
End Sub

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
    AssertArgument False, "Invalid stop trigger method"
End Select
End Function

