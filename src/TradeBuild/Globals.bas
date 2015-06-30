Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "TradeBuild27"
Private Const ModuleName                            As String = "Globals"

Public Const AttributeNameEnabled                   As String = "Enabled"
Public Const AttributeNameServiceProviderEnabled    As String = "Enabled"
Public Const AttributeNameServiceProviderProgId     As String = "ProgId"

Private Const ConfigSectionPathSeparator            As String = "/"

Public Const ConfigSectionMarketDataSources         As String = "MarketDataSources"
Public Const ConfigSectionProperties                As String = "Properties"
Public Const ConfigSectonProperty                   As String = "Property"
Public Const ConfigSectionServiceProvider           As String = "ServiceProvider"
Public Const ConfigSectionServiceProviders          As String = "ServiceProviders"
'Public Const ConfigSectionTickers                   As String = "Tickers"
Public Const ConfigSectionTradeBuild                As String = "TradeBuild"
Public Const ConfigSectionWorkspaces                As String = "Workspaces"
'Public Const ConfigSectionWorkspace                 As String = "Workspace"

Public Const ConfigSettingNoImpliedTrades           As String = "&NoImpliedTrades"
Public Const ConfigSettingNoVolumeAdjustments       As String = "&NoVolumeAdjustments"
Public Const ConfigSettingNumberOfMarketDepthRows   As String = "&NumberOfMarketDepthRows"
Public Const ConfigSettingUseExchangeTimezone       As String = "&UseExchangeTimezone"

Public Const DefaultWorkspaceName                   As String = "Default"

'Public Const ProgIdQtBarData                        As String = "QTSP27.QTHistDataServiceProvider"
'Public Const ProgIdQtRealtimeData                   As String = "QTSP27.QTRealtimeDataServiceProvider"
'Public Const ProgIdQtTickData                       As String = "QTSP27.QTTickfileServiceProvider"

Public Const ProgIdTbBarData                        As String = "TBInfoBase27.HistDataServiceProvider"
Public Const ProgIdTbContractData                   As String = "TBInfoBase27.ContractInfoSrvcProvider"
Public Const ProgIdTbOrderPersistence               As String = "TradeBuild27.OrderPersistenceSP"
Public Const ProgIdTbOrders                         As String = "TradeBuild27.OrderSimulatorSP"
Public Const ProgIdTbTickData                       As String = "TBInfoBase27.TickfileServiceProvider"

Public Const ProgIdFileTickData                     As String = "TickfileSP27.TickfileServiceProvider"

Public Const ProgIdTwsBarData                       As String = "IBTWSSP27.HistDataServiceProvider"
Public Const ProgIdTwsContractData                  As String = "IBTWSSP27.ContractInfoServiceProvider"
Public Const ProgIdTwsOrders                        As String = "IBTWSSP27.OrderSubmissionSrvcProvider"
Public Const ProgIdTwsRealtimeData                  As String = "IBTWSSP27.RealtimeDataServiceProvider"

Public Const PropertyNameRole                       As String = "Role"

Public Const PropertyNameDatabaseType               As String = "Database Type"
Public Const PropertyNameDatabaseName               As String = "Database Name"
Public Const PropertyNameDatabasePassword           As String = "Password"
Public Const PropertyNameDatabaseServer             As String = "Server"
Public Const PropertyNameDatabaseUserName           As String = "User Name"
Public Const PropertyNameDatabaseUseSynchronousWrites As String = "Use Synchronous Writes"
Public Const PropertyNameDatabaseUseSynchronousReads As String = "Use Synchronous Reads"

Public Const PropertyNameOrderPersistenceFilePath   As String = "RecoveryFilePath"

Public Const PropertyNameTfTickfileGranularity      As String = "Tickfile Granularity"
Public Const PropertyNameTfTickfilePath             As String = "Tickfile Path"

Public Const PropertyNameTwsServer                  As String = "Server"
Public Const PropertyNameTwsPort                    As String = "Port"
Public Const PropertyNameTwsClientId                As String = "Client Id"
Public Const PropertyNameTwsKeepConnection          As String = "Keep Connection"
Public Const PropertyNameTwsConnectionRetryInterval As String = "Connection Retry Interval Secs"
Public Const PropertyNameTwsLogLevel                As String = "TWS Log Level"

Public Const SrvcProviderNameRealtimeData           As String = "Realtime data"
Public Const SrvcProviderNamePrimaryContractData    As String = "Primary contract data"
Public Const SrvcProviderNameSecondaryContractData  As String = "Secondary contract data"
Public Const SrvcProviderNameHistoricalDataInput    As String = "Historical bar data retrieval"
Public Const SrvcProviderNameHistoricalDataOutput   As String = "Historical bar data storage"
Public Const SrvcProviderNameBrokerLive             As String = "Live order submission"
Public Const SrvcProviderNameBrokerSimulated        As String = "Simulated order submission"
Public Const SrvcProviderNameTickfileInput          As String = "Tickfile replay"
Public Const SrvcProviderNameTickfileOutput         As String = "Tickfile storage"
Public Const SrvcProviderNameOrderPersistence       As String = "Order persistence"

Public Const ServiceProvidersRenderer               As String = "TradeBuildUI27.SPConfigurer"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Variables
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
Const ProcName As String = "gLogger"

On Error GoTo Err

If lLogger Is Nothing Then
    Set lLogger = CreateFormattingLogger("tradebuild.log", ProjectName)
End If
Set gLogger = lLogger

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get gTracer() As Tracer
Static lTracer As Tracer
Const ProcName As String = "gTracer"
On Error GoTo Err

If lTracer Is Nothing Then Set lTracer = GetTracer("tradebuild")
Set gTracer = lTracer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function GApiNotifyCodeToString(Value As ApiNotifyCodes) As String
Select Case Value
Case ApiNotifyServiceProviderError
    GApiNotifyCodeToString = "ApiNotifyServiceProviderError"
Case ApiNotifySourceNotResponding
    GApiNotifyCodeToString = "ApiNotifySourceNotResponding"
Case ApiNotifyConnecting
    GApiNotifyCodeToString = "ApiNotifyConnecting"
Case ApiNotifyConnected
    GApiNotifyCodeToString = "ApiNotifyConnected"
Case ApiNotifyCantConnect
    GApiNotifyCodeToString = "ApiNotifyCantConnect"
Case ApiNotifyRetryingConnection
    GApiNotifyCodeToString = "ApiNotifyRetryingConnection"
Case ApiNotifyReconnecting
    GApiNotifyCodeToString = "ApiNotifyReconnecting"
Case ApiNotifyLostConnection
    GApiNotifyCodeToString = "ApiNotifyLostConnection"
Case ApiNotifyNonSpecificNotification
    GApiNotifyCodeToString = "ApiNotifyNonSpecificNotification"
Case Else
    GApiNotifyCodeToString = CStr(Value)
End Select

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
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Private Sub notifyCollectionMember( _
                ByVal pItem As Variant, _
                ByVal pSource As Object, _
                ByVal pListener As ICollectionChangeListener)
Dim ev As CollectionChangeEventData
Const ProcName As String = "notifyCollectionMember"
On Error GoTo Err

Set ev.Source = pSource
ev.changeType = CollItemAdded

gSetVariant ev.AffectedItem, pItem
pListener.Change ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
