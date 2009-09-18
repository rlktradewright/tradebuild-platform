VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiceProviders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.ICommonServiceConsumer

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                           As String = "TradeBuild26"
Private Const ModuleName                            As String = "ServiceProviders"

Private Const AccessModeReadOnly                    As String = "Read only"
Private Const AccessModeWriteOnly                   As String = "Write only"
Private Const AccessModeReadWrite                   As String = "Read write"

Private Const AttributeNameEnabled                  As String = "Enabled"
Private Const AttributeNamePropertyName             As String = "Name"
Private Const AttributeNamePropertyValue            As String = "Value"
Private Const AttributeNameServiceProviderEnabled   As String = "Enabled"
'Private Const AttributeNameServiceProviderName      As String = "Name"
Private Const AttributeNameServiceProviderProgId    As String = "ProgId"

Private Const ConfigNameProperties                  As String = "Properties"
Private Const ConfigNameProperty                    As String = "Property"
Private Const ConfigNameServiceProvider             As String = "ServiceProvider"
Private Const ConfigNameServiceProviders            As String = "ServiceProviders"

Private Const ProgIdQtBarData                       As String = "QTSP26.QTHistDataServiceProvider"
Private Const ProgIdQtRealtimeData                  As String = "QTSP26.QTRealtimeDataServiceProvider"
Private Const ProgIdQtTickData                      As String = "QTSP26.QTTickfileServiceProvider"

Private Const ProgIdTbBarData                       As String = "TBInfoBase26.HistDataServiceProvider"
Private Const ProgIdTbContractData                  As String = "TBInfoBase26.ContractInfoSrvcProvider"
Private Const ProgIdTbOrders                        As String = "TB26.OrderSubmissionSrvcProvider"
Private Const ProgIdTbTickData                      As String = "TBInfoBase26.TickfileServiceProvider"

Private Const ProgIdFileTickData                    As String = "TickfileSP26.TickfileServiceProvider"

Private Const ProgIdTwsBarData                      As String = "IBTWSSP26.HistDataServiceProvider"
Private Const ProgIdTwsContractData                 As String = "IBTWSSP26.ContractInfoServiceProvider"
Private Const ProgIdTwsOrders                       As String = "IBTWSSP26.OrderSubmissionSrvcProvider"
Private Const ProgIdTwsRealtimeData                 As String = "IBTWSSP26.RealtimeDataServiceProvider"

Private Const PropertyNameQtServer                  As String = "Server"
Private Const PropertyNameQtPort                    As String = "Port"
Private Const PropertyNameQtPassword                As String = "Password"
Private Const PropertyNameQtKeepConnection          As String = "Keep Connection"
Private Const PropertyNameQtConnectionRetryInterval As String = "Connection Retry Interval Secs"
Private Const PropertyNameQtProviderKey             As String = "Provider Key"

Private Const PropertyNameTbServer                  As String = "Server"
Private Const PropertyNameTbDbType                  As String = "Database Type"
Private Const PropertyNameTbDbName                  As String = "Database Name"
Private Const PropertyNameTbUserName                As String = "User Name"
Private Const PropertyNameTbPassword                As String = "Password"
Private Const PropertyNameTbAccessMode              As String = "Access Mode"
Private Const PropertyNameTbRole                    As String = "Role"

Private Const PropertyNameTfAccessMode              As String = "Access Mode"
Private Const PropertyNameTfTickfilePath            As String = "Tickfile Path"

Private Const PropertyNameTwsServer                 As String = "Server"
Private Const PropertyNameTwsPort                   As String = "Port"
Private Const PropertyNameTwsClientId               As String = "Client Id"
Private Const PropertyNameTwsKeepConnection         As String = "Keep Connection"
Private Const PropertyNameTwsConnectionRetryInterval    As String = "Connection Retry Interval Secs"
Private Const PropertyNameTwsProviderKey            As String = "Provider Key"
Private Const PropertyNameTwsLogLevel               As String = "TWS Log Level"

Private Const ServiceProvidersRenderer              As String = "TradeBuildUI26.SPConfigurer"

Private Const SrvcProviderNameRealtimeData          As String = "Realtime data"
Private Const SrvcProviderNamePrimaryContractData   As String = "Primary contract data"
Private Const SrvcProviderNameSecondaryContractData As String = "Secondary contract data"
Private Const SrvcProviderNameHistoricalDataInput   As String = "Historical bar data retrieval"
Private Const SrvcProviderNameHistoricalDataOutput  As String = "Historical bar data storage"
Private Const SrvcProviderNameBrokerLive            As String = "Broker - live orders"
Private Const SrvcProviderNameBrokerSimulated       As String = "Broker - simulated orders"
Private Const SrvcProviderNameTickfileInput         As String = "Tickfile replay"
Private Const SrvcProviderNameTickfileOutput        As String = "Tickfile storage"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ServiceProviderTableEntry
    name        As String
    SPTypes     As Long
    SPObject    As TradeBuildSP.ICommonServiceProvider
    SpConfig    As ConfigurationSection
End Type
'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviders()         As ServiceProviderTableEntry
Private mMaxSPIndex                 As Long

Private mChangeListeners            As Collection

Private mConfig                     As ConfigurationSection
Private mServiceProviderConfigList  As ConfigurationSection

Private mPermittedServiceProviderRoles  As ServiceProviderRoles

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mServiceProviders(9) As ServiceProviderTableEntry
mMaxSPIndex = -1

Set mChangeListeners = New Collection

mPermittedServiceProviderRoles = SPBrokerSimulated + _
                                SPHistoricalDataInput + _
                                SPPrimaryContractData + _
                                SPRealtimeData
End Sub

'@================================================================================
' ICommonServiceConsumer Interface Members
'@================================================================================

Private Sub ICommonServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal serviceProviderHandle As Long)

CommonServiceConsumerNotifyEvent eventCode, eventMessage, serviceProviderHandle
End Sub

Private Sub ICommonServiceConsumer_NotifyFatalError( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
Dim spName As String
spName = NameFromHandle(spHandle)

gTB.NotifyFatalError errorCode, _
                    errorMessage, _
                    source & IIf(spName <> "", vbCrLf & spName, "")
End Sub

Private Sub ICommonServiceConsumer_NotifyRecoverableError( _
                            ByVal errorCode As Long, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
gTB.SetServiceProviderError errorCode, _
                errorMessage, _
                NameFromHandle(spHandle)
End Sub

Private Sub ICommonServiceConsumer_RegisterServiceProvider( _
                            ByVal Handle As Long, _
                            ByVal spType As TradeBuildSP.ServiceProviderTypes)
Dim regOk As Boolean
mServiceProviders(Handle).SPTypes = mServiceProviders(Handle).SPTypes Or spType
Select Case spType
Case TradeBuildSP.ServiceProviderTypes.ContractInfo
    Dim contractInfoSP As TradeBuildSP.IContractInfoServiceProvider
    
    On Error Resume Next
    Set contractInfoSP = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not contractInfoSP Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.SecondaryContractInfo
    Dim secContractInfoSP As TradeBuildSP.IContractInfoServiceProvider
    
    On Error Resume Next
    Set secContractInfoSP = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not secContractInfoSP Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.ExchangeSimulator
    Dim exchangeSimulatorSp As TradeBuildSP.IOrderSubmissionSrvcProvider
    
    On Error Resume Next
    Set exchangeSimulatorSp = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not exchangeSimulatorSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.HistoricalData
    Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
    
    On Error Resume Next
    Set histDataSp = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not histDataSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.LiveOrderSubmission
    Dim liveOrderSp As TradeBuildSP.IOrderSubmissionSrvcProvider
    
    On Error Resume Next
    Set liveOrderSp = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not liveOrderSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.OrderPersistence
Case TradeBuildSP.ServiceProviderTypes.RealtimeData
    Dim realtimeDataSp As TradeBuildSP.IRealtimeDataServiceProvider
    
    On Error Resume Next
    Set realtimeDataSp = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not realtimeDataSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.Tickfile
    Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
    
    On Error Resume Next
    Set tickfileSp = mServiceProviders(Handle).SPObject
    On Error GoTo 0
    If Not tickfileSp Is Nothing Then regOk = True
End Select

If Not regOk Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.ServiceProviders::ICommonServiceConsumer_RegisterServiceProvider", _
                "Caller is not a type " & spType & " service provider object"
End If
    
mServiceProviders(Handle).SPTypes = mServiceProviders(Handle).SPTypes Or spType
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get IsLiveOrderSubmitterConfigured() As Boolean
IsLiveOrderSubmitterConfigured = (Not getLiveOrderSubmissionSP Is Nothing)
End Property

Friend Property Get IsAnyServiceProviderConfigured() As Boolean
IsAnyServiceProviderConfigured = mMaxSPIndex >= 0
End Property

Public Property Let PermittedServiceProviderRoles( _
                ByVal value As ServiceProviderRoles)
If IsAnyServiceProviderConfigured Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "PermittedServiceProviderRoles", _
            "Service providers are alReady configured"
End If

mPermittedServiceProviderRoles = value
End Property

Public Property Get PermittedServiceProviderRoles() As ServiceProviderRoles
PermittedServiceProviderRoles = mPermittedServiceProviderRoles
End Property

Public Property Get SPNameRealtimeData() As String
SPNameRealtimeData = SrvcProviderNameRealtimeData
End Property

Public Property Get SPNamePrimaryContractData() As String
SPNamePrimaryContractData = SrvcProviderNamePrimaryContractData
End Property

Public Property Get SPNameSecondryContractData() As String
SPNameSecondryContractData = SrvcProviderNameSecondaryContractData
End Property

Public Property Get SPNameHistoricalDataInput() As String
SPNameHistoricalDataInput = SrvcProviderNameHistoricalDataInput
End Property

Public Property Get SPNameHistoricalDataOutput() As String
SPNameHistoricalDataOutput = SrvcProviderNameHistoricalDataOutput
End Property

Public Property Get SPNameBrokerLive() As String
SPNameBrokerLive = SrvcProviderNameBrokerLive
End Property

Public Property Get SPNameBrokerSimulated() As String
SPNameBrokerSimulated = SrvcProviderNameBrokerSimulated
End Property

Public Property Get SPNameTickfileInput() As String
SPNameTickfileInput = SrvcProviderNameTickfileInput
End Property

Public Property Get SPNameTickfileOutput() As String
SPNameTickfileOutput = SrvcProviderNameTickfileOutput
End Property

Public Property Get SPProgIdQtBarData() As String
SPProgIdQtBarData = ProgIdQtBarData
End Property

Public Property Get SPProgIdQtRealtimeData() As String
SPProgIdQtRealtimeData = ProgIdQtRealtimeData
End Property

Public Property Get SPProgIdQtTickData() As String
SPProgIdQtTickData = ProgIdQtTickData
End Property

Public Property Get SPProgIdTbBarData() As String
SPProgIdTbBarData = ProgIdTbBarData
End Property

Public Property Get SPProgIdTbContractData() As String
SPProgIdTbContractData = ProgIdTbContractData
End Property

Public Property Get SPProgIdTbOrders() As String
SPProgIdTbOrders = ProgIdTbOrders
End Property

Public Property Get SPProgIdTbTickData() As String
SPProgIdTbTickData = ProgIdTbTickData
End Property

Public Property Get SPProgIdFileTickData() As String
SPProgIdFileTickData = ProgIdFileTickData
End Property

Public Property Get SPProgIdTwsBarData() As String
SPProgIdTwsBarData = ProgIdTwsBarData
End Property

Public Property Get SPProgIdTwsContractData() As String
SPProgIdTwsContractData = ProgIdTwsContractData
End Property

Public Property Get SPProgIdTwsOrders() As String
SPProgIdTwsOrders = ProgIdTwsOrders
End Property

Public Property Get SPProgIdTwsRealtimeData() As String
SPProgIdTwsRealtimeData = ProgIdTwsRealtimeData
End Property

Friend Property Get SupportedHistoricalDataPeriods() As TimePeriod()
Dim i As Long
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.HistoricalData) <> 0 Then
            Set histDataSp = mServiceProviders(i).SPObject
            SupportedHistoricalDataPeriods = histDataSp.SupportedTimePeriods
        End If
        Exit For
    End If
Next
End Property

Friend Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
SupportedInputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Replay)
End Property

Friend Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
SupportedOutputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Record)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal progId As String, _
                ByVal enabled As Boolean, _
                ByVal paramString As String, _
                Optional ByVal name As String, _
                Optional ByVal Description As String _
                ) As TradeBuildSP.ICommonServiceProvider
Dim params As parameters
Dim spindex As Long

Dim failpoint As Long
On Error GoTo Err

Set Add = CreateObject(progId)
Set params = CreateParametersFromString(paramString)

If enabled Then
    spindex = AddSP(Add, params, name)
    If spindex < 0 Then
        Set Add = Nothing
        Exit Function
    End If
End If

Set mServiceProviders(spindex).SpConfig = AddEntryToConfig(progId, _
                                                        enabled, _
                                                        params, _
                                                        name, _
                                                        Description)

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Add" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.Add value
End Sub

Friend Sub CommonServiceConsumerNotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal serviceProviderHandle As Long)

Dim ev As NotificationEvent

If serviceProviderHandle >= 0 Then Set ev.source = mServiceProviders(serviceProviderHandle).SPObject
ev.eventMessage = NameFromHandle(serviceProviderHandle) & vbCrLf & _
                eventMessage

Select Case eventCode
Case TFTickfileEmpty
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileEmpty
Case TFTickfileInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileInvalid
Case TFTickfileVersionNotSupported
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileVersionNotSupported
Case TFContractDetailsInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileContractDetailsInvalid
Case TFNoContractDetails
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileNoContractDetails
Case TFTickfileDoesntExist
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDoesntExist
Case TFCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDataSourceNotResponding
Case TFDataSourceNotResponding
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileDataSourceNotResponding
Case TFRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryingConnectionToTickfileDataSource
Case TFContractSpecifierInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyTickfileContractSpecifierInvalid
Case TFReConnectingDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToTickfileDataSource
Case TFLostConnectionToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToTickfileDataSource
Case TFConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToTickfileDataSource
Case CIConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToContractDataSource
Case CIContractSpecifierInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
Case CICantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectContractDataSource
Case CIDisconnectedFromDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromContractDataSource
Case CILostConnectionToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToContractDataSource
Case CIReConnectingDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyReConnectingContractDataSource
Case CIRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryConnectContractDataSource
Case HDCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectHistoricDataSource
Case HDRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryingConnectionToHistoricDataSource
Case HDLostConnectionToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToHistoricDataSource
Case HDReConnectingDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToHistoricDataSource
Case HDRequestInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantWriteToHistoricDataStore
Case HDDisconnectedFromDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromHistoricDataSource
Case HDRequestFailed
    ev.eventCode = ApiNotifyCodes.ApiNotifyHistoricDataRequestFailed
Case HDConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToHistoricDataSource
Case RTCantConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectRealtimeDataSource
Case RTRetryConnectDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource
Case RTLostConnectionToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToRealtimeDataSource
Case RTReConnectingDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyReconnectingToRealtimeDataSource
Case RTRequestInvalid
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
Case RTRequestFailed
    ev.eventCode = ApiNotifyCodes.ApiNotifyRealtimeDataRequestFailed
Case RTMarketDepthRequestFailed
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
Case RTDisconnectedFromDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromRealtimeDataSource
Case RTDataSourceNotResponding
    ev.eventCode = ApiNotifyCodes.ApiNotifyRealtimeDataSourceNotResponding
Case RTConnectedToDataSource
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToRealtimeDataSource
Case LOCantConnectToBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyCantConnectToBroker
Case LORetryConnectToBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyRetryConnectToBroker
Case LOLostConnectionToBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyLostConnectionToBroker
Case LOReConnectingToBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyReConnectingToBroker
Case LOUnknownContract
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
Case LONotUniqueContract
    ev.eventCode = ApiNotifyCodes.ApiNotifyInvalidRequest
Case LODisconnectedFromBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyDisconnectedFromBroker
Case LOConnectedToBroker
    ev.eventCode = ApiNotifyCodes.ApiNotifyConnectedToBroker
Case Else
    ev.eventCode = ApiNotifyCodes.ApiNotifyNonSpecificNotification
End Select
gTB.notify ev

End Sub

Friend Sub ConfigureDefaultServiceProviders( _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String)

If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRealtimeData Then
    Add "IBTWSSP26.RealtimeDataServiceProvider", _
        True, _
        "Server=" & twsServer & _
        ";Port=" & twsPort & _
        ";Client Id=" & clientId & _
        ";Connection Retry Interval Secs=" & connectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPPrimaryContractData Then
    Add "IBTWSSP26.ContractInfoServiceProvider", _
        True, _
        "Server=" & twsServer & _
        ";Port=" & twsPort & _
        ";Client Id=" & clientId & _
        ";Connection Retry Interval Secs=" & connectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPSecondaryContractData Then
    ' not relevant in default scenario
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataInput Then
    Add "IBTWSSP26.HistDataServiceProvider", _
        True, _
        "Server=" & twsServer & _
        ";Port=" & twsPort & _
        ";Client Id=" & clientId & _
        ";Connection Retry Interval Secs=" & connectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataOutput Then
    ' no sensible default for this
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerLive Then
    Add "IBTWSSP26.OrderSubmissionSrvcProvider", _
        True, _
        "Server=" & twsServer & _
        ";Port=" & twsPort & _
        ";Client Id=" & clientId & _
        ";Connection Retry Interval Secs=" & connectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerSimulated Then
    AddSP New OrderSimulatorSP, Nothing, ""
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileInput Then
    Add "TickfileSP26.TickfileServiceProvider", _
        True, _
        "Access Mode=Read only" & _
        ";Tickfile Path=" & tickfilesPath
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileOutput Then
    Add "TickfileSP26.TickfileServiceProvider", _
        True, _
        "Access Mode=Write only" & _
        ";Tickfile Path=" & tickfilesPath
End If
End Sub

Friend Function CreateHistoricalDataReader( _
                ByVal dataConsumer As TradeBuildSP.IBarDataConsumer, _
                ByVal serviceConsumer As TradeBuildSP.IBarDataInputServiceConsumer _
                ) As TradeBuildSP.IBarDataReader
                            
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
Dim histDataSPReader As TradeBuildSP.IBarDataReader
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.HistoricalData) <> 0 Then
            Set histDataSp = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

If histDataSp Is Nothing Then Exit Function

Set histDataSPReader = histDataSp.CreateBarDataReader(dataConsumer, _
                                                    serviceConsumer)

Set CreateHistoricalDataReader = histDataSPReader

End Function
                            
Friend Function CreateHistoricalDataWriter( _
                ByVal serviceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                ByVal pContract As Contract, _
                ByVal historicDataPath As String _
                ) As HistoricDataWriter
                            
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
Dim histDataSPWriter As TradeBuildSP.IBarDataWriter
Dim histDataWriter As HistoricDataWriter
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.HistoricalData) <> 0 Then
            Set histDataSp = mServiceProviders(i).SPObject
            If histDataSp.Supports(TradeBuildSP.HistoricDataServiceProviderCapabilities.HistDataStore) _
            Then
                Exit For
            Else
                Set histDataSp = Nothing
            End If
        End If
    End If
Next

If histDataSp Is Nothing Then Exit Function

Set histDataSPWriter = histDataSp.CreateBarDataWriter(serviceConsumer, _
                                                    pContract, _
                                                    historicDataPath)

Set histDataWriter = New HistoricDataWriter
histDataWriter.ServiceProviderWriter = histDataSPWriter
Set CreateHistoricalDataWriter = histDataWriter

End Function

Friend Function CreateLiveOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract _
                ) As TradeBuildSP.IOrderSubmitter
                            
Dim lOrderSimulatorSp As TradeBuildSP.IOrderSubmissionSrvcProvider
Dim orderSubmitter As TradeBuildSP.IOrderSubmitter
            
Set lOrderSimulatorSp = getLiveOrderSubmissionSP
If lOrderSimulatorSp Is Nothing Then Exit Function

Set orderSubmitter = lOrderSimulatorSp.CreateOrderSubmitter(serviceConsumer, pContract)

Set CreateLiveOrderSubmitter = orderSubmitter

                            
End Function

Friend Function CreateRealtimeDataReader( _
                ByVal pTicker As Ticker, _
                ByVal serviceConsumer As TradeBuildSP.IRealtimeDataInputServiceConsumer) As TradeBuildSP.IRealtimeDataReader
Dim realtimeSp As TradeBuildSP.IRealtimeDataServiceProvider
Dim realtimeSPReader As TradeBuildSP.IRealtimeDataReader
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.RealtimeData) <> 0 Then
            Set realtimeSp = mServiceProviders(i).SPObject
        End If
    End If
Next

If realtimeSp Is Nothing Then Exit Function

Set realtimeSPReader = realtimeSp.CreateRealtimeDataReader(pTicker, _
                                                    serviceConsumer)

Set CreateRealtimeDataReader = realtimeSPReader
End Function

Friend Function CreateSimulatedOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract) As TradeBuildSP.IOrderSubmitter
                            
Dim lOrderSimulatorSp As TradeBuildSP.IOrderSubmissionSrvcProvider
Dim orderSubmitter As TradeBuildSP.IOrderSubmitter
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.ExchangeSimulator) <> 0 Then
            Set lOrderSimulatorSp = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

If lOrderSimulatorSp Is Nothing Then
    ' no order simulator sp has been Added, so we will resort
    ' to using the default built-in simulator
    Set lOrderSimulatorSp = New OrderSimulatorSP
    AddSP lOrderSimulatorSp, Nothing, ""
End If

Set orderSubmitter = lOrderSimulatorSp.CreateOrderSubmitter(serviceConsumer, pContract)

Set CreateSimulatedOrderSubmitter = orderSubmitter

                            
End Function

Friend Function CreateTickfileReader( _
                ByVal pTicker As Ticker, _
                ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal inputTickfileFormatID As String) As TickfileReader
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPReader As TradeBuildSP.ITickfileReader
Dim tfReader As TickfileReader
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.Tickfile) <> 0 Then
            Set tickfileSp = mServiceProviders(i).SPObject
            If tickfileSp.Supports(TradeBuildSP.TickfileServiceProviderCapabilities.Replay, _
                                    inputTickfileFormatID) _
            Then
                Exit For
            Else
                Set tickfileSp = Nothing
            End If
        End If
    End If
Next

If tickfileSp Is Nothing Then Exit Function

Set tfReader = New TickfileReader
tfReader.dataConsumer = pTicker
Set tickfileSPReader = tickfileSp.CreateTickfileReader(tfReader, _
                                                    serviceConsumer, _
                                                    inputTickfileFormatID)

tfReader.ServiceProviderReader = tickfileSPReader

Set CreateTickfileReader = tfReader
End Function

Friend Function CreateTickfileWriter( _
                ByVal serviceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal pContract As Contract, _
                ByVal outputTickfileFormatID As String, _
                ByVal outputTickfilePath As String) As TickfileWriter
                            
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPWriter As TradeBuildSP.ITickfileWriter
Dim tfWriter As TickfileWriter
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.Tickfile) <> 0 Then
            Set tickfileSp = mServiceProviders(i).SPObject
            If tickfileSp.Supports(TradeBuildSP.TickfileServiceProviderCapabilities.Record, _
                                    outputTickfileFormatID) _
            Then
                Exit For
            Else
                Set tickfileSp = Nothing
            End If
        End If
    End If
Next

If tickfileSp Is Nothing Then Exit Function

Set tfWriter = New TickfileWriter
Set tickfileSPWriter = tickfileSp.CreateTickfileWriter(serviceConsumer, _
                                                    pContract, _
                                                    outputTickfileFormatID, _
                                                    outputTickfilePath)

tfWriter.ServiceProviderWriter = tickfileSPWriter

Set CreateTickfileWriter = tfWriter

End Function

Friend Function IsSupportedHistoricalDataPeriod( _
                ByRef period As TimePeriod) As Boolean
Dim i As Long
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.HistoricalData) <> 0 Then
            Set histDataSp = mServiceProviders(i).SPObject
            IsSupportedHistoricalDataPeriod = histDataSp.IsSupportedTimePeriod(period)
            Exit For
        End If
    End If
Next
End Function

Friend Function loadContracts( _
                ByVal serviceConsumer As IContractInfoServiceConsumer, _
                ByVal pContractsBuilder As ContractsBuilder) As Boolean
Dim i As Long
Dim contractInfoSP As TradeBuildSP.IContractInfoServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.ContractInfo) <> 0 Then
            Set contractInfoSP = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

loadContracts = True

If contractInfoSP Is Nothing Then
    ' there doesn't seem to be a primary contract service provider configured,
    ' so let's try the secondary
    gLog pMsg:="No primary contract info service provider configured - trying secondary", pProcName:="loadContracts", pProjName:=ProjectName, pModName:=ModuleName
    If Not loadContractsFromSecondarySP(-1, serviceConsumer, pContractsBuilder) Then
        gLog pMsg:="No secondary contract info service provider configured - can't load contracts", pProcName:="loadContracts", pProjName:=ProjectName, pModName:=ModuleName
        loadContracts = False
    End If
Else
    contractInfoSP.RequestContractDetails serviceConsumer, pContractsBuilder
End If

End Function

Friend Function loadContractsFromSecondarySP( _
                            ByVal prevSPIndex As Long, _
                            ByVal serviceConsumer As IContractInfoServiceConsumer, _
                            ByVal pContractsBuilder As ContractsBuilder _
                            ) As Boolean
Dim i As Long
Dim contractInfoSP As TradeBuildSP.IContractInfoServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.SecondaryContractInfo) <> 0 Then
            If i = prevSPIndex Then
                ' we've alReady tried the secondary, so nothing more we can do
            Else
                Set contractInfoSP = mServiceProviders(i).SPObject
            End If
            Exit For
        End If
    End If
Next

If contractInfoSP Is Nothing Then Exit Function

contractInfoSP.RequestContractDetails serviceConsumer, pContractsBuilder
loadContractsFromSecondarySP = True
End Function

Public Sub LoadServiceProviderConfiguration( _
                ByVal config As ConfigurationSection)
Dim failpoint As Long
On Error GoTo Err

Set mConfig = config

RemoveAll

On Error Resume Next
Set mServiceProviderConfigList = mConfig.GetConfigurationSection(ConfigNameServiceProviders)
On Error GoTo 0

If Not mServiceProviderConfigList Is Nothing Then
    loadConfiguration
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "LoadServiceProviderConfiguration" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Sub

Friend Function NameFromHandle( _
                ByVal Handle As Long) As String
If Handle < 0 Then Exit Function
NameFromHandle = mServiceProviders(Handle).name
End Function

Public Sub Remove( _
                ByVal sp As TradeBuildSP.ICommonServiceProvider)
Dim i As Long

If sp Is Nothing Then Exit Sub

For i = 0 To mMaxSPIndex
    If mServiceProviders(i).SPObject Is sp Then
        mServiceProviders(i).SPObject.Terminate
        RemoveSP i
        Exit For
    End If
Next
End Sub

Public Sub RemoveAll()
Dim i As Long
Dim sp As TradeBuildSP.ICommonServiceProvider

gTB.StopAllTickers

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        Set sp = mServiceProviders(i).SPObject
        sp.Terminate
        RemoveSP i
    End If
Next

mMaxSPIndex = -1

End Sub

Public Sub RemoveCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next
End Sub

Public Sub SetDefaultSPConfig( _
                ByVal configdata As ConfigurationSection, _
                Optional ByVal twsServer As String, _
                Optional ByVal twsPort As Long = 7496, _
                Optional ByVal clientId As Long = -1, _
                Optional ByVal connectionRetryIntervalSecs As Long = 60, _
                Optional ByVal twsLogLevel As String = "ERROR", _
                Optional ByVal tickfilesPath As String)
Dim currSPsList As ConfigurationSection
Dim currSP As ConfigurationSection
Dim currProps As ConfigurationSection

On Error Resume Next
Set currSPsList = configdata.GetConfigurationSection(ConfigNameServiceProviders)
On Error GoTo 0

If Not currSPsList Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "setDefaultServiceProviders", _
            "Service provider list is not empty"
End If

Set currSPsList = configdata.AddConfigurationSection(ConfigNameServiceProviders, , ServiceProvidersRenderer)

If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRealtimeData Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameRealtimeData & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTwsRealtimeData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTwsServer, twsServer
    setProperty currProps, PropertyNameTwsPort, twsPort
    setProperty currProps, PropertyNameTwsClientId, clientId
    setProperty currProps, PropertyNameTwsKeepConnection, "True"
    setProperty currProps, PropertyNameTwsConnectionRetryInterval, connectionRetryIntervalSecs
    setProperty currProps, PropertyNameTwsLogLevel, twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPPrimaryContractData Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNamePrimaryContractData & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTwsContractData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTwsServer, twsServer
    setProperty currProps, PropertyNameTwsPort, twsPort
    setProperty currProps, PropertyNameTwsClientId, clientId
    setProperty currProps, PropertyNameTwsKeepConnection, "True"
    setProperty currProps, PropertyNameTwsConnectionRetryInterval, connectionRetryIntervalSecs
    setProperty currProps, PropertyNameTwsLogLevel, twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPSecondaryContractData Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameSecondaryContractData & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTwsContractData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTwsServer, twsServer
    setProperty currProps, PropertyNameTwsPort, twsPort
    setProperty currProps, PropertyNameTwsClientId, clientId
    setProperty currProps, PropertyNameTwsKeepConnection, "True"
    setProperty currProps, PropertyNameTwsConnectionRetryInterval, connectionRetryIntervalSecs
    setProperty currProps, PropertyNameTwsLogLevel, twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataInput Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameHistoricalDataInput & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTwsBarData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTwsServer, twsServer
    setProperty currProps, PropertyNameTwsPort, twsPort
    setProperty currProps, PropertyNameTwsClientId, clientId
    setProperty currProps, PropertyNameTwsKeepConnection, "True"
    setProperty currProps, PropertyNameTwsConnectionRetryInterval, connectionRetryIntervalSecs
    setProperty currProps, PropertyNameTwsLogLevel, twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPHistoricalDataOutput Then
    ' no sensible default for this
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerLive Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameBrokerLive & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTwsOrders

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTwsServer, twsServer
    setProperty currProps, PropertyNameTwsPort, twsPort
    setProperty currProps, PropertyNameTwsClientId, clientId
    setProperty currProps, PropertyNameTwsKeepConnection, "True"
    setProperty currProps, PropertyNameTwsConnectionRetryInterval, connectionRetryIntervalSecs
    setProperty currProps, PropertyNameTwsLogLevel, twsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPBrokerSimulated Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameBrokerSimulated & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdTbOrders

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileInput Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameTickfileInput & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdFileTickData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTbAccessMode, AccessModeReadOnly
    setProperty currProps, PropertyNameTfTickfilePath, tickfilesPath
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPTickfileOutput Then
    Set currSP = currSPsList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                    "(" & SrvcProviderNameTickfileOutput & ")")
    currSP.setAttribute AttributeNameServiceProviderEnabled, "True"
    currSP.setAttribute AttributeNameServiceProviderProgId, ProgIdFileTickData

    Set currProps = currSP.AddConfigurationSection(ConfigNameProperties)
    setProperty currProps, PropertyNameTbAccessMode, AccessModeWriteOnly
    setProperty currProps, PropertyNameTfTickfilePath, tickfilesPath
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddEntryToConfig( _
            ByVal progId As String, _
            ByVal enabled As Boolean, _
            ByVal params As parameters, _
            ByVal name As String, _
            ByVal Description As String _
            ) As ConfigurationSection
Dim sp As ConfigurationSection
Dim props As ConfigurationSection
Dim prop As ConfigurationSection
Dim param As Parameter

If mConfig Is Nothing Then Exit Function

If name = "" Then name = GenerateGUIDString

If mServiceProviderConfigList Is Nothing Then
    Set mServiceProviderConfigList = mConfig.AddConfigurationSection(ConfigNameServiceProviders, ServiceProvidersRenderer)
End If

Set sp = mServiceProviderConfigList.AddConfigurationSection(ConfigNameServiceProvider, _
                                                            Description)
Set sp = mServiceProviderConfigList.AddConfigurationSection(ConfigNameServiceProvider & _
                                                            "(" & name & ")", _
                                                            Description)
sp.setAttribute AttributeNameServiceProviderEnabled, CStr(enabled)
sp.setAttribute AttributeNameServiceProviderProgId, progId

Set props = sp.AddConfigurationSection(ConfigNameProperties)

For Each param In params
    Set prop = props.AddConfigurationSection(ConfigNameProperty)
    prop.setAttribute AttributeNamePropertyName, param.name
    prop.setAttribute AttributeNamePropertyValue, param.value
Next

Set AddEntryToConfig = sp
End Function

Private Function AddSP( _
            ByVal serviceProvider As TradeBuildSP.ICommonServiceProvider, _
            ByVal params As parameters, _
            ByVal name As String _
            ) As Long
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If name <> "" Then serviceProvider.name = name

' we don't reuse Removed entries to ensure that service providers Added later are
' always later in the table than those Added earlier
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If mServiceProviders(i).SPObject Is serviceProvider Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.ServiceProviders::Add", _
                    "This Service Provider object has alReady been Added"
        End If
        If mServiceProviders(i).name = serviceProvider.name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.ServiceProviders::Add", _
                    "A Service Provider with this name has alReady been Added"
        End If
    End If
Next

If mMaxSPIndex = UBound(mServiceProviders) Then
    ReDim Preserve mServiceProviders(UBound(mServiceProviders) + 10) As ServiceProviderTableEntry
End If
mMaxSPIndex = mMaxSPIndex + 1

mServiceProviders(mMaxSPIndex).name = serviceProvider.name
Set mServiceProviders(mMaxSPIndex).SPObject = serviceProvider

serviceProvider.Link Me, mMaxSPIndex, params

If mServiceProviders(i).SPObject Is Nothing Then
    ' a fatal error has occurred while linking the SP - eg, invalid parameter
    ' has been passed
    AddSP = -1
Else
    fireChange CollItemAdded, serviceProvider
    AddSP = mMaxSPIndex
End If

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "AddSP" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As TradeBuildSP.ICommonServiceProvider)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub

Private Function getLiveOrderSubmissionSP() As TradeBuildSP.IOrderSubmissionSrvcProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.LiveOrderSubmission) <> 0 Then
            Set getLiveOrderSubmissionSP = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

End Function

Private Function GetSupportedTickfileFormats( _
                            ByVal capabilities As TradeBuildSP.TickfileServiceProviderCapabilities _
                            ) As TickfileFormatSpecifier()
Dim i As Long
Dim j As Long
Dim k As Long
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim supportedFormats() As TickfileFormatSpecifier
ReDim allSupportedFormats(9) As TickfileFormatSpecifier

k = -1
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.Tickfile) <> 0 Then
            Set tickfileSp = mServiceProviders(i).SPObject
            supportedFormats = tickfileSp.supportedFormats
            For j = 0 To UBound(supportedFormats)
                If (supportedFormats(j).capabilities And capabilities) <> 0 Then
                    k = k + 1
                    If k > UBound(allSupportedFormats) Then
                        ReDim Preserve allSupportedFormats(UBound(allSupportedFormats) + 10) As TickfileFormatSpecifier
                    End If
                    allSupportedFormats(k) = supportedFormats(j)
                End If
            Next
        End If
    End If
Next

If k >= 0 Then
    ReDim Preserve allSupportedFormats(k) As TickfileFormatSpecifier
    GetSupportedTickfileFormats = allSupportedFormats
Else
    Erase allSupportedFormats
End If
End Function

Private Sub loadConfiguration()
Dim sp As ConfigurationSection
Dim spName As String
Dim spProgId As String
Dim props As ConfigurationSection
Dim prop As ConfigurationSection
Dim propName As String
Dim propValue As String
Dim spObj As Object
Dim spindex As Long
Dim params As parameters

Dim failpoint As Long
On Error GoTo Err

For Each sp In mServiceProviderConfigList
    If CBool(sp.getAttribute(AttributeNameEnabled, "False")) Then
        
        spName = sp.InstanceQualifier
        
        spProgId = sp.getAttribute(AttributeNameServiceProviderProgId)
        If spProgId = "" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                       ProjectName & "." & ModuleName & ":" & "loadConfiguration", _
                                       "Service provider ProgId not specified"
        
        Set params = New parameters
        
        On Error Resume Next
        Set props = Nothing
        Set props = sp.GetConfigurationSection(ConfigNameProperties)
        On Error GoTo Err
        
        If Not props Is Nothing Then
            For Each prop In props
                If prop.InstanceQualifier = "" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                                            ProjectName & "." & ModuleName & ":" & "loadConfiguration", _
                                                            "Property name not specified"
                'propValue = prop.getAttribute(AttributeNamePropertyValue, "")
                params.SetParameterValue prop.InstanceQualifier, prop.value
            Next
        End If
        
        On Error Resume Next
        If spProgId = ProgIdTbOrders Then
            ' this is the built-in order simulator
            Set spObj = New OrderSimulatorSP
        Else
            Set spObj = CreateObject(spProgId)
        End If
        On Error GoTo Err
        If spObj Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                       ProjectName & "." & ModuleName & ":" & "loadConfiguration", _
                                       "Can't create service provider with progId " & spProgId
        
        spindex = AddSP(spObj, params, spName)
        If spindex >= 0 Then
            Set mServiceProviders(spindex).SpConfig = sp
        Else
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & "loadConfiguration", _
                    "Can't Add service provider " & spName
        End If
    End If
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.number
Dim errSource As String: errSource = IIf(Err.source <> "", Err.source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "loadConfiguration" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub RemoveSP( _
                ByVal Handle As Long, _
                Optional ByVal RemoveFromConfig As Boolean)
Dim sp As ICommonServiceProvider

If Handle < 0 Then Exit Sub

Set sp = mServiceProviders(Handle).SPObject

If RemoveFromConfig And Not mServiceProviders(Handle).SpConfig Is Nothing Then
    mServiceProviderConfigList.RemoveConfigurationSection mServiceProviders(Handle).SpConfig
End If

Set mServiceProviders(Handle).SPObject = Nothing
Set mServiceProviders(Handle).SpConfig = Nothing
mServiceProviders(Handle).SPTypes = 0
mServiceProviders(Handle).name = ""
fireChange CollItemRemoved, sp
End Sub

Private Sub setProperty( _
                ByVal props As ConfigurationSection, _
                ByVal name As String, _
                ByVal value As String)
'Dim prop As ConfigurationSection
props.SetSetting "." & ConfigNameProperty & "(" & name & ")", value
'prop.setAttribute AttributeNamePropertyName, name
'prop.setAttribute AttributeNamePropertyValue, value
End Sub



