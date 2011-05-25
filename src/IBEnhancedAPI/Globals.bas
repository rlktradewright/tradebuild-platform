Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const ProjectName                        As String = "IBTWSSP27"
Private Const ModuleName                        As String = "Globals"

Public Const InvalidEnumValue                   As String = "*ERR*"
Public Const NullIndex                          As Long = -1

Public Const MaxLong                            As Long = &H7FFFFFFF
Public Const OneMicrosecond                     As Double = 1# / 86400000000#
Public Const OneMinute                          As Double = 1# / 1440#
Public Const OneSecond                          As Double = 1# / 86400#

Public Const NumDaysInWeek                      As Long = 5
Public Const NumDaysInMonth                     As Long = 22
Public Const NumDaysInYear                      As Long = 260
Public Const NumMonthsInYear                    As Long = 12

Public Const ContractInfoSPName                 As String = "IB Tws Contract Info Service Provider"
Public Const HistoricDataSPName                 As String = "IB Tws Historic Data Service Provider"
Public Const RealtimeDataSPName                 As String = "IB Tws Realtime Data Service Provider"
Public Const OrderSubmissionSPName              As String = "IB Tws Order Submission Service Provider"

Public Const ProviderKey                        As String = "Tws"

Public Const ParamNameClientId                  As String = "Client Id"
Public Const ParamNameConnectionRetryIntervalSecs As String = "Connection Retry Interval Secs"
Public Const ParamNameKeepConnection            As String = "Keep Connection"
Public Const ParamNamePort                      As String = "Port"
Public Const ParamNameProviderKey               As String = "Provider Key"
Public Const ParamNameRole                      As String = "Role"
Public Const ParamNameServer                    As String = "Server"
Public Const ParamNameTwsLogLevel               As String = "Tws Log Level"
Public Const ParamNameDisableRequestPacing      As String = "Disable Request Pacing"

Public Const TwsLogLevelDetailString            As String = "Detail"
Public Const TwsLogLevelErrorString             As String = "Error"
Public Const TwsLogLevelInformationString       As String = "Information"
Public Const TwsLogLevelSystemString            As String = "System"
Public Const TwsLogLevelWarningString           As String = "Warning"

'================================================================================
' Enums
'================================================================================

Public Enum ConnectionStates
    ConnNotConnected
    ConnConnecting
    ConnReSynching
    ConnConnected
End Enum

Public Enum FADataTypes
    FAGroups = 1
    FaProfile
    FAAccountAliases
End Enum

Public Enum TwsSocketInMsgTypes
    TICK_PRICE = 1
    TICK_SIZE = 2
    ORDER_STATUS = 3
    ERR_MSG = 4
    OPEN_ORDER = 5
    ACCT_Value = 6
    PORTFOLIO_Value = 7
    ACCT_UPDATE_TIME = 8
    NEXT_VALID_ID = 9
    CONTRACT_DATA = 10
    EXECUTION_DATA = 11
    MARKET_DEPTH = 12
    MARKET_DEPTH_L2 = 13
    NEWS_BULLETINS = 14
    MANAGED_ACCTS = 15
    RECEIVE_FA = 16
    HISTORICAL_DATA = 17
    BOND_CONTRACT_DATA = 18
    SCANNER_PARAMETERS = 19
    SCANNER_DATA = 20
    TICK_OPTION_COMPUTATION = 21
    TICK_GENERIC = 45
    TICK_STRING = 46
    TICK_EFP = 47
    CURRENT_TIME = 49
    REAL_TIME_BARS = 50
    FUNDAMENTAL_DATA = 51
    CONTRACT_DATA_END = 52
    OPEN_ORDER_END = 53
    ACCT_DOWNLOAD_END = 54
    EXECUTION_DATA_END = 55
    DELTA_NEUTRAL_VALIDATION = 56
    TICK_SNAPSHOT_END = 57
    MAX_SOCKET_INMSG
End Enum

Public Enum TwsSocketOutMsgTypes
    REQ_MKT_DATA = 1
    CANCEL_MKT_DATA = 2
    PLACE_ORDER = 3
    CANCEL_ORDER = 4
    REQ_OPEN_ORDERS = 5
    REQ_ACCT_DATA = 6
    REQ_EXECUTIONS = 7
    REQ_IDS = 8
    REQ_CONTRACT_DATA = 9
    REQ_MKT_DEPTH = 10
    CANCEL_MKT_DEPTH = 11
    REQ_NEWS_BULLETINS = 12
    CANCEL_NEWS_BULLETINS = 13
    SET_SERVER_LOGLEVEL = 14
    REQ_AUTO_OPEN_ORDERS = 15
    REQ_ALL_OPEN_ORDERS = 16
    REQ_MANAGED_ACCTS = 17
    REQ_FA = 18
    REPLACE_FA = 19
    REQ_HISTORICAL_DATA = 20
    EXERCISE_OPTIONS = 21
    REQ_SCANNER_SUBSCRIPTION = 22
    CANCEL_SCANNER_SUBSCRIPTION = 23
    REQ_SCANNER_PARAMETERS = 24
    CANCEL_HISTORICAL_DATA = 25
    REQ_CURRENT_TIME = 49
    REQ_REAL_TIME_BARS = 50
    CANCEL_REAL_TIME_BARS = 51
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Global variables
'================================================================================

'================================================================================
' Private variables
'================================================================================

Private mLogger As FormattingLogger

'================================================================================
' Properties
'================================================================================

Public Property Get gLogger() As FormattingLogger
If mLogger Is Nothing Then Set mLogger = CreateFormattingLogger("tradebuild.log.serviceprovider.ibtwssp", ProjectName)
Set gLogger = mLogger
End Property

'================================================================================
' Methods
'================================================================================

Public Function gContractSpecToTwsContract(ByVal pContractSpecifier As ContractSpecifier) As TwsContract
Const ProcName As String = "gContractSpecToTwsContract"
On Error GoTo Err

Dim lComboLeg As comboLeg
Dim lTwsComboLeg As TwsComboLeg

Set gContractSpecToTwsContract = New TwsContract

With gContractSpecToTwsContract
    .CurrencyCode = pContractSpecifier.CurrencyCode
    .Exchange = pContractSpecifier.Exchange
    .Expiry = pContractSpecifier.Expiry
    .IncludeExpired = True
    .LocalSymbol = pContractSpecifier.LocalSymbol
    .OptRight = gOptionRightToTwsOptRight(pContractSpecifier.Right)
    .Sectype = gSecTypeToTwsSecType(pContractSpecifier.Sectype)
    .Strike = pContractSpecifier.Strike
    .Symbol = pContractSpecifier.Symbol
    If Not pContractSpecifier.ComboLegs Is Nothing Then
        For Each lComboLeg In pContractSpecifier.ComboLegs
            Set lTwsComboLeg = New TwsComboLeg
            With lTwsComboLeg
                .Action = IIf(lComboLeg.IsBuyLeg, TwsOrderActions.TwsOrderActionBuy, TwsOrderActions.TwsOrderActionSell)
                Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Combo contracts not supported"
                ' Need to fix this: the problem is that we need to do a contract details
                ' request to discover the contact id for the combo leg
            End With
        Next
    End If
End With

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gContractToTwsContractDetails(ByVal pContract As Contract) As TwsContractDetails
Const ProcName As String = "gContractToTwsContractDetails"
On Error GoTo Err

Dim lContract As TwsContract
Dim lContractDetails As TwsContractDetails

If pContract.Specifier.Sectype = SecTypeCombo Then Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Combo contracts not supported"

Set lContractDetails = New TwsContractDetails
Set lContract = gContractSpecToTwsContract(pContract.Specifier)

With lContractDetails
    .Summary = lContract
    .MinTick = pContract.TickSize
    .TimeZoneId = gStandardTimezoneNameToTwsTimeZoneName(pContract.TimeZone.StandardName)
End With

Set gContractToTwsContractDetails = lContractDetails

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gCreateContractDetailsFetcher( _
                ByVal pClient As TwsAPI) As ContractDetailsFetcher
Const ProcName As String = "CreateContractDetailsFetcher"
On Error GoTo Err

Set gCreateContractDetailsFetcher = New ContractDetailsFetcher
gCreateContractDetailsFetcher.Initialise getContractDetailsRequester(pClient)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gCreateHistoricalDataFetcher( _
                ByVal pClient As TwsAPI) As HistoricalDataFetcher
Const ProcName As String = "gCreateHistoricalDataFetcher"
On Error GoTo Err

Dim lHistDataRequester As HistDataRequester

Set lHistDataRequester = getHistDataRequester(pClient)
If lHistDataRequester Is Nothing Then Set lHistDataRequester = setupHistDataRequester(pClient)

Set gCreateHistoricalDataFetcher = New HistoricalDataFetcher
gCreateHistoricalDataFetcher.Initialise lHistDataRequester, getContractDetailsRequester(pClient)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gCreateMarketDataSource( _
                ByVal pClient As TwsAPI, _
                ByVal pContract As Contract) As MarketDataSource
Const ProcName As String = "gCreateMarketDataSource"
On Error GoTo Err

Dim lMarketDataRequester As MarketDataRequester

If Not pClient.Tws.MarketDataConsumer Is Nothing Then
    If Not TypeOf pClient.Tws.MarketDataConsumer Is MarketDataRequester Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Tws is already configured with an incompatible IMarketDataConsumer"
    Set lMarketDataRequester = pClient.Tws.MarketDataConsumer
Else
    Set lMarketDataRequester = New MarketDataRequester
    lMarketDataRequester.Initialise pClient
    pClient.Tws.MarketDataConsumer = lMarketDataRequester
    pClient.Tws.MarketDepthConsumer = lMarketDataRequester
End If

Dim lMarketDataSource As New MarketDataSource
lMarketDataSource.Initialise lMarketDataRequester, pContract, getContractDetailsRequester(pClient)

Set gCreateMarketDataSource = lMarketDataSource

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gCreateOrderSubmitter( _
                ByVal pClient As TwsAPI, _
                ByVal pContractSpec As ContractSpecifier, _
                ByVal pOrderSubmissionListener As IOrderSubmissionListener) As OrderSubmitter
Const ProcName As String = "gCreateOrderSubmitter"
On Error GoTo Err

Dim lOrderPlacer As OrderPlacer

If Not pClient.Tws.OrderInfoConsumer Is Nothing Then
    If Not TypeOf pClient.Tws.OrderInfoConsumer Is OrderPlacer Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Tws is already configured with an incompatible IOrderInfoConsumer"
    Set lOrderPlacer = pClient.Tws.OrderInfoConsumer
Else
    Set lOrderPlacer = New OrderPlacer
    lOrderPlacer.Initialise pClient
    pClient.Tws.OrderInfoConsumer = lOrderPlacer
End If

Set gCreateOrderSubmitter = New OrderSubmitter
gCreateOrderSubmitter.Initialise pClient, lOrderPlacer, pContractSpec, pOrderSubmissionListener, getContractDetailsRequester(pClient)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub gDisableHistoricalDataRequestPacing(ByVal pClient As TwsAPI)
Const ProcName As String = "gDisableHistoricalDataRequestPacing"
On Error GoTo Err

If getHistDataRequester(pClient) Is Nothing Then setupHistDataRequester(pClient).DisableHistoricalDataRequestPacing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

' format is yyyymmdd [hh:mm:ss [timezone]]
Public Function gGetDate( _
                ByRef pDateString As String, _
                Optional ByRef pTimezoneName As String) As Date
Const ProcName As String = "gGetDate"
On Error GoTo Err

If Len(pDateString) = 8 Then
    gGetDate = CDate(Left$(pDateString, 4) & "/" & _
                    Mid$(pDateString, 5, 2) & "/" & _
                    Mid$(pDateString, 7, 2))
ElseIf Len(pDateString) >= 17 Then
    gGetDate = CDate(Left$(pDateString, 4) & "/" & _
                        Mid$(pDateString, 5, 2) & "/" & _
                        Mid$(pDateString, 7, 2) & " " & _
                        Mid$(pDateString, 10, 8))
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid date string format"
End If

If IsMissing(pTimezoneName) Then Exit Function

If Len(pDateString) > 17 Then
    pTimezoneName = Trim$(Right$(pDateString, Len(pDateString) - 17))
Else
    pTimezoneName = ""
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gGetTwsAPIInstance(ByVal pServer As String, _
                                    ByVal pPort As String, _
                                    ByVal pClientId As Long, _
                                    ByVal pConnectionRetryIntervalSecs As Long, _
                                    ByVal pTwsLogLevel As TwsLogLevels) As TwsAPI
Const ProcName As String = "gGetTwsAPIInstance"
On Error GoTo Err

Set gGetTwsAPIInstance = gGetTws(pServer, pPort).GetAPI(pClientId)
gGetTwsAPIInstance.ConnectionRetryIntervalSecs = pConnectionRetryIntervalSecs
gGetTwsAPIInstance.TwsLogLevel = pTwsLogLevel

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gHistDataCapabilities() As Long
Const ProcName As String = "gHistDataCapabilities"
On Error GoTo Err

gHistDataCapabilities = 0

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
Const ProcName As String = "gHistDataSupports"
On Error GoTo Err

gHistDataSupports = (gHistDataCapabilities And capabilities)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gInputMessageIdToString( _
                ByVal msgId As TwsSocketInMsgTypes) As String
Const ProcName As String = "gInputMessageIdToString"
On Error GoTo Err

Select Case msgId
Case TICK_PRICE
    gInputMessageIdToString = "TICK_PRICE"
Case TICK_SIZE
    gInputMessageIdToString = "TICK_SIZE"
Case ORDER_STATUS
    gInputMessageIdToString = "ORDER_STATUS"
Case ERR_MSG
    gInputMessageIdToString = "ERR_MSG"
Case OPEN_ORDER
    gInputMessageIdToString = "OPEN_ORDER"
Case ACCT_Value
    gInputMessageIdToString = "ACCT_Value"
Case PORTFOLIO_Value
    gInputMessageIdToString = "PORTFOLIO_Value"
Case ACCT_UPDATE_TIME
    gInputMessageIdToString = "ACCT_UPDATE_TIME"
Case NEXT_VALID_ID
    gInputMessageIdToString = "NEXT_VALID_ID"
Case CONTRACT_DATA
    gInputMessageIdToString = "CONTRACT_DATA"
Case EXECUTION_DATA
    gInputMessageIdToString = "EXECUTION_DATA"
Case MARKET_DEPTH
    gInputMessageIdToString = "MARKET_DEPTH"
Case MARKET_DEPTH_L2
    gInputMessageIdToString = "MARKET_DEPTH_L2"
Case NEWS_BULLETINS
    gInputMessageIdToString = "NEWS_BULLETINS"
Case MANAGED_ACCTS
    gInputMessageIdToString = "MANAGED_ACCTS"
Case RECEIVE_FA
    gInputMessageIdToString = "RECEIVE_FA"
Case HISTORICAL_DATA
    gInputMessageIdToString = "HISTORICAL_DATA"
Case BOND_CONTRACT_DATA
    gInputMessageIdToString = "BOND_CONTRACT_DATA"
Case SCANNER_PARAMETERS
    gInputMessageIdToString = "SCANNER_PARAMETERS"
Case SCANNER_DATA
    gInputMessageIdToString = "SCANNER_DATA"
Case TICK_OPTION_COMPUTATION
    gInputMessageIdToString = "TICK_OPTION_COMPUTATION"
Case TICK_GENERIC
    gInputMessageIdToString = "TICK_GENERIC"
Case TICK_STRING
    gInputMessageIdToString = "TICK_STRING"
Case TICK_EFP
    gInputMessageIdToString = "TICK_EFP"
Case CURRENT_TIME
    gInputMessageIdToString = "CURRENT_TIME"
Case REAL_TIME_BARS
    gInputMessageIdToString = "REAL_TIME_BARS"
Case FUNDAMENTAL_DATA
    gInputMessageIdToString = "FUNDAMENTAL_DATA"
Case CONTRACT_DATA_END
    gInputMessageIdToString = "CONTRACT_DATA_END"
Case OPEN_ORDER_END
    gInputMessageIdToString = "OPEN_ORDER_END"
Case ACCT_DOWNLOAD_END
    gInputMessageIdToString = "ACCT_DOWNLOAD_END"
Case EXECUTION_DATA_END
    gInputMessageIdToString = "EXECUTION_DATA_END"
Case DELTA_NEUTRAL_VALIDATION
    gInputMessageIdToString = "DELTA_NEUTRAL_VALIDATION"
Case TICK_SNAPSHOT_END
    gInputMessageIdToString = "TICK_SNAPSHOT_END"
Case Else
    gInputMessageIdToString = "?????"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub gLog(ByRef pMsg As String, _
                ByRef pModName As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
Const ProcName As String = "gLog"
On Error GoTo Err


gLogger.Log pMsg, pProcName, pModName, pLogLevel, pMsgQualifier

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function gOptionRightToTwsOptRight(ByVal Value As OptionRights) As TwsOptionRights
Select Case Value
Case OptionRights.OptCall
    gOptionRightToTwsOptRight = TwsOptRightCall
Case OptionRights.OptPut
    gOptionRightToTwsOptRight = TwsOptRightPut
Case OptionRights.OptNone
    gOptionRightToTwsOptRight = TwsOptRightNone
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Option Right"
End Select
End Function

Public Function gOrderActionFromString(ByVal Value As String) As OrderActions
Select Case UCase$(Value)
Case ""
    gOrderActionFromString = OrderActionNone
Case "BUY"
    gOrderActionFromString = OrderActionBuy
Case "SELL"
    gOrderActionFromString = OrderActionSell
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gOrderActionToTwsOrderAction(ByVal Value As OrderActions) As TwsOrderActions
Select Case Value
Case OrderActions.OrderActionBuy
    gOrderActionToTwsOrderAction = TwsOrderActionBuy
Case OrderActions.OrderActionSell
    gOrderActionToTwsOrderAction = TwsOrderActionSell
Case OrderActions.OrderActionNone
    gOrderActionToTwsOrderAction = TwsOrderActionNone
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid OrderAction"
End Select
End Function

Public Function gOrderStatusFromString(ByVal Value As String) As OrderStatuses
Select Case UCase$(Value)
Case "CREATED"
    gOrderStatusFromString = OrderStatusCreated
Case "REJECTED", "INACTIVE"
    gOrderStatusFromString = OrderStatusRejected
Case "PENDINGSUBMIT"
    gOrderStatusFromString = OrderStatusPendingSubmit
Case "PRESUBMITTED"
    gOrderStatusFromString = OrderStatusPreSubmitted
Case "SUBMITTED"
    gOrderStatusFromString = OrderStatusSubmitted
Case "PENDINGCANCEL"
    gOrderStatusFromString = OrderStatusCancelling
Case "CANCELLED"
    gOrderStatusFromString = OrderStatusCancelled
Case "FILLED"
    gOrderStatusFromString = OrderStatusFilled
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid order status: " & Value
End Select
End Function

Public Function gOrderTIFFromString(ByVal Value As String) As OrderTIFs
Select Case UCase$(Value)
Case ""
    gOrderTIFFromString = OrderTIFNone
Case "DAY"
    gOrderTIFFromString = OrderTIFDay
Case "GTC"
    gOrderTIFFromString = OrderTIFGoodTillCancelled
Case "IOC"
    gOrderTIFFromString = OrderTIFImmediateOrCancel
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order TIF"
End Select
End Function

Public Function gOrderTIFToString(ByVal Value As OrderTIFs) As String
Select Case Value
Case OrderTIFs.OrderTIFDay
    gOrderTIFToString = "DAY"
Case OrderTIFs.OrderTIFGoodTillCancelled
    gOrderTIFToString = "GTC"
Case OrderTIFs.OrderTIFImmediateOrCancel
    gOrderTIFToString = "IOC"
Case OrderTIFs.OrderTIFNone
    gOrderTIFToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order TIF"
End Select
End Function

Public Function gOrderTIFToTwsOrderTIF(ByVal Value As OrderTIFs) As TwsOrderTIFs
Select Case Value
Case OrderTIFNone
    gOrderTIFToTwsOrderTIF = TwsOrderTIFNone
Case OrderTIFDay
    gOrderTIFToTwsOrderTIF = TwsOrderTIFDay
Case OrderTIFGoodTillCancelled
    gOrderTIFToTwsOrderTIF = TwsOrderTIFGoodTillCancelled
Case OrderTIFImmediateOrCancel
    gOrderTIFToTwsOrderTIF = TwsOrderTIFImmediateOrCancel
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid OrderTIF"
End Select
End Function

Public Function gOrderToTwsOrder(ByVal pOrder As IOrder) As TwsOrder
Const ProcName As String = "gOrderToTwsOrder"
On Error GoTo Err

Set gOrderToTwsOrder = New TwsOrder
With pOrder
    gOrderToTwsOrder.Action = gOrderActionToTwsOrderAction(.Action)
    gOrderToTwsOrder.AllOrNone = .AllOrNone
    gOrderToTwsOrder.BlockOrder = .BlockOrder
    gOrderToTwsOrder.OrderId = .BrokerId
    gOrderToTwsOrder.DiscretionaryAmt = .DiscretionaryAmount
    gOrderToTwsOrder.DisplaySize = .DisplaySize
    gOrderToTwsOrder.ETradeOnly = .ETradeOnly
    gOrderToTwsOrder.FirmQuoteOnly = .FirmQuoteOnly
    If .GoodAfterTime <> 0 Then gOrderToTwsOrder.GoodAfterTime = Format(.GoodAfterTime, "yyyymmdd hh:nn:ss") & IIf(.GoodAfterTimeTZ <> "", " " & .GoodAfterTimeTZ, "")
    If .GoodTillDate <> 0 Then gOrderToTwsOrder.GoodTillDate = Format(.GoodTillDate, "yyyymmdd hh:nn:ss") & IIf(.GoodTillDateTZ <> "", " " & .GoodTillDateTZ, "")
    gOrderToTwsOrder.Hidden = .Hidden
    gOrderToTwsOrder.OutsideRth = .IgnoreRegularTradingHours
    gOrderToTwsOrder.LmtPrice = .LimitPrice
    gOrderToTwsOrder.MinQty = .MinimumQuantity
    gOrderToTwsOrder.NbboPriceCap = .NbboPriceCap
    gOrderToTwsOrder.OrderType = gOrderTypeToTwsOrderType(.OrderType)
    gOrderToTwsOrder.Origin = .Origin
    gOrderToTwsOrder.OrderRef = .OriginatorRef
    gOrderToTwsOrder.OverridePercentageConstraints = .OverrideConstraints
    gOrderToTwsOrder.TotalQuantity = .Quantity
    gOrderToTwsOrder.SettlingFirm = .SettlingFirm
    gOrderToTwsOrder.TriggerMethod = gStopTriggerMethodToTwsTriggerMethod(.StopTriggerMethod)
    gOrderToTwsOrder.SweepToFill = .SweepToFill
    gOrderToTwsOrder.Tif = gOrderTIFToTwsOrderTIF(.TimeInForce)
    gOrderToTwsOrder.AuxPrice = .TriggerPrice
End With

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gOrderTypeToTwsOrderType(ByVal Value As OrderTypes) As TwsOrderTypes
Const ProcName As String = "gOrderTypeToTwsOrderType"
On Error GoTo Err

Select Case Value
Case OrderTypes.OrderTypeMarket
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarket
Case OrderTypes.OrderTypeMarketOnClose
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case OrderTypes.OrderTypeLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimit
Case OrderTypes.OrderTypeLimitOnClose
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case OrderTypes.OrderTypePeggedToMarket
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case OrderTypes.OrderTypeStop
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStop
Case OrderTypes.OrderTypeStopLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeStopLimit
Case OrderTypes.OrderTypeTrail
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrail
Case OrderTypes.OrderTypeRelative
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeRelative
Case OrderTypes.OrderTypeVWAP
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeVWAP
Case OrderTypes.OrderTypeMarketToLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case OrderTypes.OrderTypeQuote
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeQuote
Case OrderTypes.OrderTypeAdjust
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeAdjust
Case OrderTypes.OrderTypeAlert
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeAlert
Case OrderTypes.OrderTypeLimitIfTouched
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case OrderTypes.OrderTypeMarketIfTouched
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case OrderTypes.OrderTypeTrailLimit
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeTrailLimit
Case OrderTypes.OrderTypeMarketWithProtection
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case OrderTypes.OrderTypeMarketOnOpen
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case OrderTypes.OrderTypeLimitOnOpen
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case OrderTypes.OrderTypePeggedToPrimary
    gOrderTypeToTwsOrderType = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid OrderType"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gOutputMessageIdToString( _
                ByVal msgId As TwsSocketOutMsgTypes) As String
Const ProcName As String = "gOutputMessageIdToString"
On Error GoTo Err

Select Case msgId
Case REQ_MKT_DATA
    gOutputMessageIdToString = "REQ_MKT_DATA"
Case CANCEL_MKT_DATA
    gOutputMessageIdToString = "CANCEL_MKT_DATA"
Case PLACE_ORDER
    gOutputMessageIdToString = "PLACE_ORDER"
Case CANCEL_ORDER
    gOutputMessageIdToString = "CANCEL_ORDER"
Case REQ_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_OPEN_ORDERS"
Case REQ_ACCT_DATA
    gOutputMessageIdToString = "REQ_ACCT_DATA"
Case REQ_EXECUTIONS
    gOutputMessageIdToString = "REQ_EXECUTIONS"
Case REQ_IDS
    gOutputMessageIdToString = "REQ_IDS"
Case REQ_CONTRACT_DATA
    gOutputMessageIdToString = "REQ_CONTRACT_DATA"
Case REQ_MKT_DEPTH
    gOutputMessageIdToString = "REQ_MKT_DEPTH"
Case CANCEL_MKT_DEPTH
    gOutputMessageIdToString = "CANCEL_MKT_DEPTH"
Case REQ_NEWS_BULLETINS
    gOutputMessageIdToString = "REQ_NEWS_BULLETINS"
Case CANCEL_NEWS_BULLETINS
    gOutputMessageIdToString = "CANCEL_NEWS_BULLETINS"
Case SET_SERVER_LOGLEVEL
    gOutputMessageIdToString = "SET_SERVER_LOGLEVEL"
Case REQ_AUTO_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_AUTO_OPEN_ORDERS"
Case REQ_ALL_OPEN_ORDERS
    gOutputMessageIdToString = "REQ_ALL_OPEN_ORDERS"
Case REQ_MANAGED_ACCTS
    gOutputMessageIdToString = "REQ_MANAGED_ACCTS"
Case REQ_FA
    gOutputMessageIdToString = "REQ_FA"
Case REPLACE_FA
    gOutputMessageIdToString = "REPLACE_FA"
Case REQ_HISTORICAL_DATA
    gOutputMessageIdToString = "REQ_HISTORICAL_DATA"
Case EXERCISE_OPTIONS
    gOutputMessageIdToString = "EXERCISE_OPTIONS"
Case REQ_SCANNER_SUBSCRIPTION
    gOutputMessageIdToString = "REQ_SCANNER_SUBSCRIPTION"
Case CANCEL_SCANNER_SUBSCRIPTION
    gOutputMessageIdToString = "CANCEL_SCANNER_SUBSCRIPTION"
Case REQ_SCANNER_PARAMETERS
    gOutputMessageIdToString = "REQ_SCANNER_PARAMETERS"
Case CANCEL_HISTORICAL_DATA
    gOutputMessageIdToString = "CANCEL_HISTORICAL_DATA"
Case REQ_CURRENT_TIME
    gOutputMessageIdToString = "REQ_CURRENT_TIME"
Case REQ_REAL_TIME_BARS
    gOutputMessageIdToString = "REQ_REAL_TIME_BARS"
Case CANCEL_REAL_TIME_BARS
    gOutputMessageIdToString = "CANCEL_REAL_TIME_BARS"
Case Else
    gOutputMessageIdToString = "?????"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gParseClientId( _
                Value As String) As Long
Const ProcName As String = "gParseClientId"

On Error GoTo Err

If Value = "" Then
    gParseClientId = -1
ElseIf Not IsInteger(Value) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Client Id' parameter: Value must be an integer"
Else
    gParseClientId = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gParseConnectionRetryInterval( _
                Value As String) As Long
Const ProcName As String = "gParseConnectionRetryInterval"

On Error GoTo Err

If Value = "" Then
    gParseConnectionRetryInterval = 0
ElseIf Not IsInteger(Value, 0) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Connection Retry Interval Secs' parameter: Value must be an integer >= 0"
Else
    gParseConnectionRetryInterval = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gParseKeepConnection( _
                Value As String) As Boolean
Const ProcName As String = "gParseKeepConnection"
On Error GoTo Err
If Value = "" Then
    gParseKeepConnection = False
Else
    gParseKeepConnection = CBool(Value)
End If
Exit Function

Err:
Err.Raise ErrorCodes.ErrIllegalArgumentException, _
        ProjectName & "." & ModuleName & ":" & ProcName, _
        "Invalid 'Keep Connection' parameter: Value must be 'true' or 'false'"
End Function

Public Function gParsePort( _
                Value As String) As Long
Const ProcName As String = "gParsePort"

On Error GoTo Err

If Value = "" Then
    gParsePort = 7496
ElseIf Not IsInteger(Value, 1024, 65535) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Port' parameter: Value must be an integer >= 1024 and <=65535"
Else
    gParsePort = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gParseRole( _
                Value As String) As String
Const ProcName As String = "gParseRole"



On Error GoTo Err

Select Case UCase$(Value)
Case "", "P", "PR", "PRIM", "PRIMARY"
    gParseRole = "PRIMARY"
Case "S", "SEC", "SECOND", "SECONDARY"
    gParseRole = "SECONDARY"
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Role' parameter: Value must be one of 'P', 'PR', 'PRIM', 'PRIMARY', 'S', 'SEC', 'SECOND', or 'SECONDARY'"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gParseTwsLogLevel( _
                Value As String) As TwsLogLevels
Const ProcName As String = "gParseTwsLogLevel"

On Error GoTo Err

If Value = "" Then
    gParseTwsLogLevel = TwsLogLevelError
Else
    gParseTwsLogLevel = gTwsLogLevelFromString(Value)
End If
Exit Function

Err:
If Err.number = ErrorCodes.ErrIllegalArgumentException Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid 'Tws Log Level' parameter: Value must be one of " & _
            TwsLogLevelSystemString & ", " & _
            TwsLogLevelErrorString & ", " & _
            TwsLogLevelWarningString & ", " & _
            TwsLogLevelInformationString & " or " & _
            TwsLogLevelDetailString
End If

gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gRealtimeDataCapabilities() As Long
Const ProcName As String = "gRealtimeDataCapabilities"
gRealtimeDataCapabilities = RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition
End Function

Public Function gRealtimeDataSupports(ByVal capabilities As Long) As Boolean
Const ProcName As String = "gRealtimeDataSupports"
gRealtimeDataSupports = (gRealtimeDataCapabilities And capabilities)
End Function

Public Sub gReleaseTwsAPIInstance( _
                ByVal pTwsAPI As TwsAPI, _
                ByVal pForceDisconnect As Boolean)
Const ProcName As String = "gReleaseTwsAPIInstance"
On Error GoTo Err

pTwsAPI.Tws.ReleaseAPI pTwsAPI, pForceDisconnect

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function gRoundTimeToSecond( _
                ByVal Timestamp As Date) As Date
Const ProcName As String = "gRoundTimeToSecond"
gRoundTimeToSecond = Int((Timestamp + (499 / 86400000)) * 86400) / 86400 + 1 / 86400000000#
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gSecTypeToTwsSecType(ByVal Value As SecurityTypes) As TwsSecTypes
Select Case UCase$(Value)
Case SecurityTypes.SecTypeStock
    gSecTypeToTwsSecType = TwsSecTypeStock
Case SecurityTypes.SecTypeFuture
    gSecTypeToTwsSecType = TwsSecTypeFuture
Case SecurityTypes.SecTypeOption
    gSecTypeToTwsSecType = TwsSecTypeOption
Case SecurityTypes.SecTypeFuturesOption
    gSecTypeToTwsSecType = TwsSecTypeFuturesOption
Case SecurityTypes.SecTypeCash
    gSecTypeToTwsSecType = TwsSecTypeCash
Case SecurityTypes.SecTypeCombo
    gSecTypeToTwsSecType = TwsSecTypeCombo
Case SecurityTypes.SecTypeIndex
    gSecTypeToTwsSecType = TwsSecTypeIndex
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Security Type"
End Select
End Function

Public Function gSocketInMsgTypeToString( _
                ByVal Value As TwsSocketInMsgTypes) As String
Const ProcName As String = "gSocketInMsgTypeToString"
On Error GoTo Err

Select Case Value
Case TICK_PRICE
    gSocketInMsgTypeToString = "Tick price          "
Case TICK_SIZE
    gSocketInMsgTypeToString = "Tick Size           "
Case ORDER_STATUS
    gSocketInMsgTypeToString = "Order status        "
Case ERR_MSG
    gSocketInMsgTypeToString = "Error message       "
Case OPEN_ORDER
    gSocketInMsgTypeToString = "Open Order          "
Case ACCT_Value
    gSocketInMsgTypeToString = "Account Value       "
Case PORTFOLIO_Value
    gSocketInMsgTypeToString = "Portfolio Value     "
Case ACCT_UPDATE_TIME
    gSocketInMsgTypeToString = "Account update Time "
Case NEXT_VALID_ID
    gSocketInMsgTypeToString = "Next valid id       "
Case CONTRACT_DATA
    gSocketInMsgTypeToString = "Contract data       "
Case EXECUTION_DATA
    gSocketInMsgTypeToString = "Execution data      "
Case MARKET_DEPTH
    gSocketInMsgTypeToString = "Market depth        "
Case MARKET_DEPTH_L2
    gSocketInMsgTypeToString = "Market depth L2     "
Case NEWS_BULLETINS
    gSocketInMsgTypeToString = "New bulletin        "
Case MANAGED_ACCTS
    gSocketInMsgTypeToString = "Managed accounts    "
Case RECEIVE_FA
    gSocketInMsgTypeToString = "Receive FA          "
Case HISTORICAL_DATA
    gSocketInMsgTypeToString = "Historical data     "
Case BOND_CONTRACT_DATA
    gSocketInMsgTypeToString = "Bond contract data  "
Case SCANNER_PARAMETERS
    gSocketInMsgTypeToString = "scanner parameters  "
Case SCANNER_DATA
    gSocketInMsgTypeToString = "Scanner data        "
Case TICK_OPTION_COMPUTATION
    gSocketInMsgTypeToString = "Option computation  "
Case TICK_GENERIC
    gSocketInMsgTypeToString = "Generic             "
Case TICK_STRING
    gSocketInMsgTypeToString = "String              "
Case TICK_EFP
    gSocketInMsgTypeToString = "EFP                 "
Case CURRENT_TIME
    gSocketInMsgTypeToString = "Current Time        "
Case REAL_TIME_BARS
    gSocketInMsgTypeToString = "Realtime bar        "
Case FUNDAMENTAL_DATA
    gSocketInMsgTypeToString = "Fundamental data    "
Case CONTRACT_DATA_END
    gSocketInMsgTypeToString = "Contract data end   "
Case OPEN_ORDER_END
    gSocketInMsgTypeToString = "Open Order end      "
Case ACCT_DOWNLOAD_END
    gSocketInMsgTypeToString = "Account download end"
Case EXECUTION_DATA_END
    gSocketInMsgTypeToString = "Execution data end  "
Case DELTA_NEUTRAL_VALIDATION
    gSocketInMsgTypeToString = "Delta neutral validn"
Case TICK_SNAPSHOT_END
    gSocketInMsgTypeToString = "Tick snapshot end   "
Case Else
    gSocketInMsgTypeToString = "Msg type " & Format(Value, "00                  ")
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Public Function gStandardTimezoneNameToTwsTimeZoneName(ByVal pTimezoneName As String) As String
Const ProcName As String = "gStandardTimezoneNameToTwsTimeZoneName"
On Error GoTo Err

Select Case pTimezoneName
Case ""
    gStandardTimezoneNameToTwsTimeZoneName = ""
Case "AUS Eastern Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "AET"
Case "Central Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "CTT"
Case "GMT Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "GMT"
Case "Eastern Standard Time"
    gStandardTimezoneNameToTwsTimeZoneName = "EST"
Case Else
    gLog "Unrecognised timezone", ModuleName, ProcName, pTimezoneName
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gStopTriggerMethodToTwsTriggerMethod(ByVal Value As StopTriggerMethods) As TwsStopTriggerMethods
Select Case Value
Case StopTriggerDefault
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDefault
Case StopTriggerDoubleBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerBidAsk
Case StopTriggerLast
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLast
Case StopTriggerDoubleLast
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerDoubleLast
Case StopTriggerBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerBidAsk
Case StopTriggerLastOrBidAsk
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerLastOrBidAsk
Case StopTriggerMidPoint
    gStopTriggerMethodToTwsTriggerMethod = TwsStopTriggerMidPoint
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid StopTriggerMethod"
End Select
End Function
                
Public Function gTruncateTimeToNextMinute(ByVal Timestamp As Date) As Date
Const ProcName As String = "gTruncateTimeToNextMinute"
On Error GoTo Err

gTruncateTimeToNextMinute = Int((Timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTruncateTimeToMinute(ByVal Timestamp As Date) As Date
Const ProcName As String = "gTruncateTimeToMinute"
On Error GoTo Err

gTruncateTimeToMinute = Int((Timestamp + OneMicrosecond) / OneMinute) * OneMinute

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTwsContractDetailsToContract(ByVal pTwsContractDetails As TwsContractDetails) As Contract
Const ProcName As String = "gTwsContractDetailsToContract"
On Error GoTo Err

Dim lBuilder As ContractBuilder

With pTwsContractDetails
    With .Summary
        Set lBuilder = CreateContractBuilder(CreateContractSpecifier(.LocalSymbol, .Symbol, .Exchange, gTwsSecTypeToSecType(.Sectype), .CurrencyCode, .Expiry, .Strike, gTwsOptionRightToOptionRight(.OptRight)))
        If .Expiry <> "" Then
            lBuilder.ExpiryDate = CDate(Left$(.Expiry, 4) & "/" & _
                                                Mid$(.Expiry, 5, 2) & "/" & _
                                                Right$(.Expiry, 2))
        End If
        lBuilder.Multiplier = .Multiplier
    End With
    lBuilder.Description = .LongName
    lBuilder.TickSize = .MinTick
    lBuilder.TimeZone = GetTimeZone(gTwsTimezoneNameToStandardTimeZoneName(.TimeZoneId))

End With

Set gTwsContractDetailsToContract = lBuilder.Contract

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTwsLogLevelFromString( _
                ByVal Value As String) As TwsLogLevels
Const ProcName As String = "gTwsLogLevelFromString"

On Error GoTo Err

Select Case UCase$(Value)
Case UCase$(TwsLogLevelDetailString)
    gTwsLogLevelFromString = TwsLogLevelDetail
Case UCase$(TwsLogLevelErrorString)
    gTwsLogLevelFromString = TwsLogLevelError
Case UCase$(TwsLogLevelInformationString)
    gTwsLogLevelFromString = TwsLogLevelInformation
Case UCase$(TwsLogLevelSystemString)
    gTwsLogLevelFromString = TwsLogLevelSystem
Case UCase$(TwsLogLevelWarningString)
    gTwsLogLevelFromString = TwsLogLevelWarning
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Log Level"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTwsOptionRightFromString(ByVal Value As String) As TwsOptionRights
Select Case UCase$(Value)
Case ""
    gTwsOptionRightFromString = TwsOptRightNone
Case "CALL", "C"
    gTwsOptionRightFromString = TwsOptRightCall
Case "PUT", "P"
    gTwsOptionRightFromString = TwsOptRightPut
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Option Right"
End Select
End Function

Public Function gTwsOptionRightToOptionRight(ByVal Value As TwsOptionRights) As OptionRights
Select Case Value
Case TwsOptRightNone
    gTwsOptionRightToOptionRight = OptionRights.OptNone
Case TwsOptRightCall
    gTwsOptionRightToOptionRight = OptionRights.OptCall
Case TwsOptRightPut
    gTwsOptionRightToOptionRight = OptionRights.OptPut
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Option Right"
End Select
End Function

Public Function gTwsOptionRightToString(ByVal Value As TwsOptionRights) As String
Select Case Value
Case TwsOptRightNone
    gTwsOptionRightToString = ""
Case TwsOptRightCall
    gTwsOptionRightToString = "Call"
Case TwsOptRightPut
    gTwsOptionRightToString = "Put"
Case Else
    gTwsOptionRightToString = InvalidEnumValue
End Select
End Function

Public Function gTwsOrderActionFromString(ByVal Value As String) As TwsOrderActions
Select Case UCase$(Value)
Case ""
    gTwsOrderActionFromString = TwsOrderActionNone
Case "BUY"
    gTwsOrderActionFromString = TwsOrderActionBuy
Case "SELL"
    gTwsOrderActionFromString = TwsOrderActionSell
Case "SSHORT"
    gTwsOrderActionFromString = TwsOrderActionSellShort
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderActionToOrderAction(ByVal Value As TwsOrderActions) As OrderActions
Select Case Value
Case TwsOrderActionBuy
    gTwsOrderActionToOrderAction = OrderActionBuy
Case TwsOrderActionSell
    gTwsOrderActionToOrderAction = OrderActionSell
Case TwsOrderActionNone
    gTwsOrderActionToOrderAction = OrderActionNone
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderActionToString(ByVal Value As TwsOrderActions) As String
Select Case Value
Case TwsOrderActionBuy
    gTwsOrderActionToString = "BUY"
Case TwsOrderActionSell
    gTwsOrderActionToString = "SELL"
Case TwsOrderActionSellShort
    gTwsOrderActionToString = "SSHORT"
Case TwsOrderActionNone
    gTwsOrderActionToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Order Action"
End Select
End Function

Public Function gTwsOrderTIFFromString(ByVal Value As String) As TwsOrderTIFs
Select Case UCase$(Value)
Case ""
    gTwsOrderTIFFromString = TwsOrderTIFNone
Case "DAY"
    gTwsOrderTIFFromString = TwsOrderTIFDay
Case "GTC"
    gTwsOrderTIFFromString = TwsOrderTIFGoodTillCancelled
Case "IOC"
    gTwsOrderTIFFromString = TwsOrderTIFImmediateOrCancel
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function gTwsOrderTIFToString(ByVal Value As TwsOrderTIFs) As String
Select Case Value
Case TwsOrderTIFs.TwsOrderTIFDay
    gTwsOrderTIFToString = "DAY"
Case TwsOrderTIFs.TwsOrderTIFGoodTillCancelled
    gTwsOrderTIFToString = "GTC"
Case TwsOrderTIFs.TwsOrderTIFImmediateOrCancel
    gTwsOrderTIFToString = "IOC"
Case TwsOrderTIFs.TwsOrderTIFNone
    gTwsOrderTIFToString = ""
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Tws Order TIF"
End Select
End Function

Public Function gTwsOrderTypeFromString(ByVal Value As String) As TwsOrderTypes
Select Case UCase$(Value)
Case "MKT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarket
Case "MKTCLS"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case "LMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimit
Case "LMTCLS"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case "PEGMKT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToMarket
Case "STP"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStop
Case "STPLMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeStopLimit
Case "TRAIL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrail
Case "REL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeRelative
Case "VWAP"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVWAP
Case "MTL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketToLimit
Case "RFQ"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeQuote
Case "ADJUST"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAdjust
Case "ALERT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeAlert
Case "LIT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitIfTouched
Case "MIT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketIfTouched
Case "TRAILLMT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeTrailLimit
Case "MKTPROT"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketWithProtection
Case "MOO"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnOpen
Case "MOC"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeMarketOnClose
Case "LOO"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnOpen
Case "LOC"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeLimitOnClose
Case "PEGPRI"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypePeggedToPrimary
Case "VOL"
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeVol
Case Else
    gTwsOrderTypeFromString = TwsOrderTypes.TwsOrderTypeNone
End Select
End Function

Public Function gTwsOrderTypeToOrderType(ByVal Value As TwsOrderTypes) As OrderTypes
Const ProcName As String = "gTwsOrderTypeToOrderType"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeMarket
    gTwsOrderTypeToOrderType = OrderTypeMarket
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    gTwsOrderTypeToOrderType = OrderTypeMarketOnClose
Case TwsOrderTypes.TwsOrderTypeLimit
    gTwsOrderTypeToOrderType = OrderTypeLimit
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    gTwsOrderTypeToOrderType = OrderTypeLimitOnClose
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    gTwsOrderTypeToOrderType = OrderTypePeggedToMarket
Case TwsOrderTypes.TwsOrderTypeStop
    gTwsOrderTypeToOrderType = OrderTypeStop
Case TwsOrderTypes.TwsOrderTypeStopLimit
    gTwsOrderTypeToOrderType = OrderTypeStopLimit
Case TwsOrderTypes.TwsOrderTypeTrail
    gTwsOrderTypeToOrderType = OrderTypeTrail
Case TwsOrderTypes.TwsOrderTypeRelative
    gTwsOrderTypeToOrderType = OrderTypeRelative
Case TwsOrderTypes.TwsOrderTypeVWAP
    gTwsOrderTypeToOrderType = OrderTypeVWAP
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    gTwsOrderTypeToOrderType = OrderTypeMarketToLimit
Case TwsOrderTypes.TwsOrderTypeQuote
    gTwsOrderTypeToOrderType = OrderTypeQuote
Case TwsOrderTypes.TwsOrderTypeAdjust
    gTwsOrderTypeToOrderType = OrderTypeAdjust
Case TwsOrderTypes.TwsOrderTypeAlert
    gTwsOrderTypeToOrderType = OrderTypeAlert
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    gTwsOrderTypeToOrderType = OrderTypeLimitIfTouched
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    gTwsOrderTypeToOrderType = OrderTypeMarketIfTouched
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    gTwsOrderTypeToOrderType = OrderTypeTrailLimit
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    gTwsOrderTypeToOrderType = OrderTypeMarketWithProtection
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    gTwsOrderTypeToOrderType = OrderTypeMarketOnOpen
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    gTwsOrderTypeToOrderType = OrderTypeLimitOnOpen
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    gTwsOrderTypeToOrderType = OrderTypePeggedToPrimary
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Unsupported order type"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTwsOrderTypeToString(ByVal Value As TwsOrderTypes) As String
Const ProcName As String = "gTwsOrderTypeToString"
On Error GoTo Err

Select Case Value
Case TwsOrderTypes.TwsOrderTypeMarket
    gTwsOrderTypeToString = "MKT"
Case TwsOrderTypes.TwsOrderTypeMarketOnClose
    gTwsOrderTypeToString = "MKTCLS"
Case TwsOrderTypes.TwsOrderTypeLimit
    gTwsOrderTypeToString = "LMT"
Case TwsOrderTypes.TwsOrderTypeLimitOnClose
    gTwsOrderTypeToString = "LMTCLS"
Case TwsOrderTypes.TwsOrderTypePeggedToMarket
    gTwsOrderTypeToString = "PEGMKT"
Case TwsOrderTypes.TwsOrderTypeStop
    gTwsOrderTypeToString = "STP"
Case TwsOrderTypes.TwsOrderTypeStopLimit
    gTwsOrderTypeToString = "STPLMT"
Case TwsOrderTypes.TwsOrderTypeTrail
    gTwsOrderTypeToString = "TRAIL"
Case TwsOrderTypes.TwsOrderTypeRelative
    gTwsOrderTypeToString = "REL"
Case TwsOrderTypes.TwsOrderTypeVWAP
    gTwsOrderTypeToString = "VWAP"
Case TwsOrderTypes.TwsOrderTypeMarketToLimit
    gTwsOrderTypeToString = "MTL"
Case TwsOrderTypes.TwsOrderTypeQuote
    gTwsOrderTypeToString = "QUOTE"
Case TwsOrderTypes.TwsOrderTypeAdjust
    gTwsOrderTypeToString = "ADJUST"
Case TwsOrderTypes.TwsOrderTypeAlert
    gTwsOrderTypeToString = "ALERT"
Case TwsOrderTypes.TwsOrderTypeLimitIfTouched
    gTwsOrderTypeToString = "LIT"
Case TwsOrderTypes.TwsOrderTypeMarketIfTouched
    gTwsOrderTypeToString = "MIT"
Case TwsOrderTypes.TwsOrderTypeTrailLimit
    gTwsOrderTypeToString = "TRAILLMT"
Case TwsOrderTypes.TwsOrderTypeMarketWithProtection
    gTwsOrderTypeToString = "MKTPROT"
Case TwsOrderTypes.TwsOrderTypeMarketOnOpen
    gTwsOrderTypeToString = "MOO"
Case TwsOrderTypes.TwsOrderTypeLimitOnOpen
    gTwsOrderTypeToString = "LOO"
Case TwsOrderTypes.TwsOrderTypePeggedToPrimary
    gTwsOrderTypeToString = "PEGPRI"
Case TwsOrderTypes.TwsOrderTypeVol
    gTwsOrderTypeToString = "VOL"
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTwsSecTypeFromString(ByVal Value As String) As TwsSecTypes
Select Case UCase$(Value)
Case "STOCK", "STK"
    gTwsSecTypeFromString = TwsSecTypeStock
Case "FUTURE", "FUT"
    gTwsSecTypeFromString = TwsSecTypeFuture
Case "OPTION", "OPT"
    gTwsSecTypeFromString = TwsSecTypeOption
Case "FUTURES OPTION", "FOP"
    gTwsSecTypeFromString = TwsSecTypeFuturesOption
Case "CASH"
    gTwsSecTypeFromString = TwsSecTypeCash
Case "BAG", "COMBO", "CMB"
    gTwsSecTypeFromString = TwsSecTypeCombo
Case "INDEX", "IND"
    gTwsSecTypeFromString = TwsSecTypeIndex
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Security Type"
End Select
End Function

Public Function gTwsSecTypeToSecType(ByVal Value As TwsSecTypes) As SecurityTypes
Select Case Value
Case TwsSecTypeStock
    gTwsSecTypeToSecType = SecurityTypes.SecTypeStock
Case TwsSecTypeFuture
    gTwsSecTypeToSecType = SecurityTypes.SecTypeFuture
Case TwsSecTypeOption
    gTwsSecTypeToSecType = SecurityTypes.SecTypeOption
Case TwsSecTypeFuturesOption
    gTwsSecTypeToSecType = SecurityTypes.SecTypeFuturesOption
Case TwsSecTypeCash
    gTwsSecTypeToSecType = SecurityTypes.SecTypeCash
Case TwsSecTypeCombo
    gTwsSecTypeToSecType = SecurityTypes.SecTypeCombo
Case TwsSecTypeIndex
    gTwsSecTypeToSecType = SecurityTypes.SecTypeIndex
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Security Type"
End Select
End Function

Public Function gTwsSecTypeToString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    gTwsSecTypeToString = "Stock"
Case TwsSecTypeFuture
    gTwsSecTypeToString = "Future"
Case TwsSecTypeOption
    gTwsSecTypeToString = "Option"
Case TwsSecTypeFuturesOption
    gTwsSecTypeToString = "Futures Option"
Case TwsSecTypeCash
    gTwsSecTypeToString = "Cash"
Case TwsSecTypeCombo
    gTwsSecTypeToString = "Combo"
Case TwsSecTypeIndex
    gTwsSecTypeToString = "Index"
Case TwsSecTypeNone
    gTwsSecTypeToString = ""
Case Else
    gTwsSecTypeToString = InvalidEnumValue
End Select
End Function

Public Function gTwsSecTypeToShortString(ByVal Value As TwsSecTypes) As String
Select Case Value
Case TwsSecTypeStock
    gTwsSecTypeToShortString = "STK"
Case TwsSecTypeFuture
    gTwsSecTypeToShortString = "FUT"
Case TwsSecTypeOption
    gTwsSecTypeToShortString = "OPT"
Case TwsSecTypeFuturesOption
    gTwsSecTypeToShortString = "FOP"
Case TwsSecTypeCash
    gTwsSecTypeToShortString = "CASH"
Case TwsSecTypeCombo
    gTwsSecTypeToShortString = "BAG"
Case TwsSecTypeIndex
    gTwsSecTypeToShortString = "IND"
Case TwsSecTypeNone
    gTwsSecTypeToShortString = ""
Case Else
    gTwsSecTypeToShortString = InvalidEnumValue
End Select
End Function

Public Function gTwsShortSaleSlotFromString( _
                ByVal Value As String) As TwsShortSaleSlotCodes
Select Case UCase$(Value)
Case "", "N/A", "NOT APPLICABLE"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotNotApplicable
Case "B", "CB", "BROKER", "CLEARING BROKER"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotClearingBroker
Case "T", "TP", "THIRD PARTY"
    gTwsShortSaleSlotFromString = TwsShortSaleSlotThirdParty
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid Short Sale Slot"
End Select
End Function

Public Function gTwsShortSaleSlotToString( _
                ByVal Value As TwsShortSaleSlotCodes) As String
Select Case Value
Case TwsShortSaleSlotNotApplicable
    gTwsShortSaleSlotToString = "N/A"
Case TwsShortSaleSlotClearingBroker
    gTwsShortSaleSlotToString = "BROKER"
Case TwsShortSaleSlotThirdParty
    gTwsShortSaleSlotToString = "THIRD PARTY"
Case Else
    gTwsShortSaleSlotToString = InvalidEnumValue
End Select
End Function

Public Function gTwsStopTriggerMethodFromString(ByVal Value As String) As TwsStopTriggerMethods
Select Case Value
Case "Default"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDefault
Case "Double Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
Case "Last"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLast
Case "Double Last"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerDoubleLast
Case "Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerBidAsk
Case "Last or Bid/Ask"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
Case "Midpoint"
    gTwsStopTriggerMethodFromString = TwsStopTriggerMethods.TwsStopTriggerMidPoint
End Select
End Function

Public Function gTwsStopTriggerMethodToString(ByVal Value As TwsStopTriggerMethods) As String
Select Case Value
Case TwsStopTriggerMethods.TwsStopTriggerBidAsk
    gTwsStopTriggerMethodToString = "Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDefault
    gTwsStopTriggerMethodToString = "Default"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleBidAsk
    gTwsStopTriggerMethodToString = "Double Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerDoubleLast
    gTwsStopTriggerMethodToString = "Double Last"
Case TwsStopTriggerMethods.TwsStopTriggerLast
    gTwsStopTriggerMethodToString = "Last"
Case TwsStopTriggerMethods.TwsStopTriggerLastOrBidAsk
    gTwsStopTriggerMethodToString = "Last or Bid/Ask"
Case TwsStopTriggerMethods.TwsStopTriggerMidPoint
    gTwsStopTriggerMethodToString = "Midpoint"
End Select
End Function

Public Function gTwsTimezoneNameToStandardTimeZoneName(ByVal pTimeZoneId As String) As String
Const ProcName As String = "gTwsTimezoneNameToStandardTimeZoneName"
On Error GoTo Err

Select Case pTimeZoneId
Case ""
    gTwsTimezoneNameToStandardTimeZoneName = ""
Case "AET"
    gTwsTimezoneNameToStandardTimeZoneName = "AUS Eastern Standard Time"
Case "CTT"
    gTwsTimezoneNameToStandardTimeZoneName = "Central Standard Time"
Case "GMT"
    gTwsTimezoneNameToStandardTimeZoneName = "GMT Standard Time"
Case "EST"
    gTwsTimezoneNameToStandardTimeZoneName = "Eastern Standard Time"
Case Else
    gLog "Unrecognised timezone", ModuleName, ProcName, pTimeZoneId
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function getContractDetailsRequester( _
                ByVal pClient As TwsAPI) As ContractDetailsRequester
Const ProcName As String = "getContractDetailsRequester"
On Error GoTo Err

Dim lContractDetailsRequester As ContractDetailsRequester

If Not pClient.Tws.ContractDetailsConsumer Is Nothing Then
    If Not TypeOf pClient.Tws.ContractDetailsConsumer Is ContractDetailsRequester Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Tws is already configured with an incompatible IContractDetailsConsumer"
    Set lContractDetailsRequester = pClient.Tws.ContractDetailsConsumer
Else
    Set lContractDetailsRequester = New ContractDetailsRequester
    lContractDetailsRequester.Initialise pClient
    pClient.Tws.ContractDetailsConsumer = lContractDetailsRequester
End If

Set getContractDetailsRequester = lContractDetailsRequester

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getHistDataRequester( _
                ByVal pClient As TwsAPI) As HistDataRequester
Const ProcName As String = "getHistDataRequester "
On Error GoTo Err

Dim lHistDataRequester As HistDataRequester

If Not pClient.Tws.HistDataConsumer Is Nothing Then
    If Not TypeOf pClient.Tws.HistDataConsumer Is HistDataRequester Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Tws is already configured with an incompatible IHistDataConsumer"
    Set lHistDataRequester = pClient.Tws.HistDataConsumer
End If

Set getHistDataRequester = lHistDataRequester

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function setupHistDataRequester(ByVal pClient As TwsAPI) As HistDataRequester
Const ProcName As String = "setupHistDataRequester"
On Error GoTo Err

Set setupHistDataRequester = New HistDataRequester
setupHistDataRequester.Initialise pClient
pClient.Tws.HistDataConsumer = setupHistDataRequester

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function
    

