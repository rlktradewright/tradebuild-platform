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

Implements IServiceProviderManager

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged(ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                            As String = "ServiceProviders"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ServiceProviderTableEntry
    Name        As String
    SPRoles     As ServiceProviderRoles
    SPObject    As IServiceProvider
    SpConfig    As ConfigurationSection
    Params      As Parameters
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mPermittedServiceProviderRoles              As ServiceProviderRoles

Private mServiceProviders()                         As ServiceProviderTableEntry
Private mMaxSPIndex                                 As Long

Private mCollectionChangeListeners                  As Listeners

Private mConfig                                     As ConfigurationSection
Private mServiceProviderConfig                      As ConfigurationSection

Private mHistDataStoreInput                         As IHistoricalDataStore
Private mHistDataStoreOutput                        As IHistoricalDataStore
Private mMarketDataManager                          As IMarketDataManager
Private mOrderPersistenceDataStore                  As IOrderPersistenceDataStore
Private mOrderRecoveryAgent                         As IOrderRecoveryAgent
Private mOrderSubmitterFactoryLive                  As IOrderSubmitterFactory
Private mOrderSubmitterFactorySimulated             As IOrderSubmitterFactory
Private mContractStorePrimary                       As IContractStore
Private mContractStoreSecondary                     As IContractStore
Private mTickfileStoreInput                         As ITickfileStore
Private mTickfileStoreOutput                        As ITickfileStore

Private mServiceProvidersStarted                    As Boolean

Private mTradeBuildAPIRef                           As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mServiceProviders(9) As ServiceProviderTableEntry
mMaxSPIndex = -1

Set mCollectionChangeListeners = New Listeners

End Sub

'@================================================================================
' IServiceProviderManager Interface Members
'@================================================================================

Private Sub IServiceProviderManager_NotifyEvent(ByVal pEventCode As Long, ByVal pEventMessage As String, ByVal pServiceProviderHandle As Long)
Const ProcName As String = "IServiceProviderManager_NotifyEvent"
On Error GoTo Err

Dim ev As NotificationEventData

ev.EventCode = pEventCode
If pServiceProviderHandle >= 0 Then Set ev.Source = mServiceProviders(pServiceProviderHandle).SPObject
ev.EventMessage = NameFromHandle(pServiceProviderHandle) & " (" & _
                mServiceProviders(pServiceProviderHandle).SPObject.Details.EXEName & ")" & vbCrLf & _
                pEventMessage

tb.NotifyEvent ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Private Sub IServiceProviderManager_NotifyFatalError(ByVal pErrorCode As Long, ByVal pSource As String, ByVal pErrorMessage As String, ByVal pServiceProviderHandle As Long)
'Const ProcName As String = "IServiceProviderManager_NotifyFatalError"
'On Error GoTo Err
'
'Dim spName As String
'spName = NameFromHandle(pServiceProviderHandle)
'
'tb.NotifyFatalError pErrorCode, _
'                    pErrorMessage, _
'                    pSource & IIf(spName <> "", vbCrLf & spName, "")
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Private Sub IServiceProviderManager_NotifyRecoverableError(ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pServiceProviderHandle As Long)
Const ProcName As String = "IServiceProviderManager_NotifyRecoverableError"
On Error GoTo Err

tb.SetServiceProviderError pErrorCode, _
                pErrorMessage, _
                NameFromHandle(pServiceProviderHandle)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IServiceProviderManager_RegisterServiceProvider( _
                ByVal pHandle As Long, _
                ByVal pRole As ServiceProviderRoles)
Const ProcName As String = "IServiceProviderManager_RegisterServiceProvider"
On Error GoTo Err

Dim lSPObj As Object
Set lSPObj = mServiceProviders(pHandle).SPObject

checkSPHasPermittedRole pRole, pHandle
Select Case pRole
Case SPRoleRealtimeData
    If Not TypeOf lSPObj Is IRealtimeDataSP Then handleSPWithUnexpectedType "IRealtimeDataSP", pHandle
Case SPRoleContractDataPrimary
    If Not TypeOf lSPObj Is IContractSP Then handleSPWithUnexpectedType "IContractSP", pHandle
Case SPRoleContractDataSecondary
    If Not TypeOf lSPObj Is IContractSP Then handleSPWithUnexpectedType "IContractSP", pHandle
Case SPRoleHistoricalDataInput
    If Not TypeOf lSPObj Is IHistDataSP Then handleSPWithUnexpectedType "IHistDataSP", pHandle
Case SPRoleHistoricalDataOutput
    If Not TypeOf lSPObj Is IHistDataSP Then handleSPWithUnexpectedType "IHistDataSP", pHandle
Case SPRoleOrderSubmissionSimulated
    If Not TypeOf lSPObj Is IOrderSubmissionSP Then handleSPWithUnexpectedType "IOrderSubmissionSP", pHandle
Case SPRoleTickfileInput
    If Not TypeOf lSPObj Is ITickfileSP Then handleSPWithUnexpectedType "ITickfileSP", pHandle
Case SPRoleTickfileOutput
    If Not TypeOf lSPObj Is ITickfileSP Then handleSPWithUnexpectedType "ITickfileSP", pHandle
Case SPRoleOrderSubmissionLive
    If Not TypeOf lSPObj Is IOrderSubmissionSP Then handleSPWithUnexpectedType "IOrderSubmissionSP", pHandle
Case SPRoleOrderPersistence
    If Not TypeOf lSPObj Is IOrderPersistenceSP Then handleSPWithUnexpectedType "IOrderPersistenceSP", pHandle
End Select

mServiceProviders(pHandle).SPRoles = mServiceProviders(pHandle).SPRoles Or pRole

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ConfigurationSection(ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

Assert mConfig Is Nothing, "Configuration section has already been set"
Set mConfig = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ContractStorePrimary() As IContractStore
Set ContractStorePrimary = mContractStorePrimary
End Property

Friend Property Get ContractStoreSecondary() As IContractStore
Set ContractStoreSecondary = mContractStoreSecondary
End Property

Friend Property Get HistoricalDataStoreInput() As IHistoricalDataStore
Set HistoricalDataStoreInput = mHistDataStoreInput
End Property

Friend Property Get HistoricalDataStoreOutput() As IHistoricalDataStore
Set HistoricalDataStoreOutput = mHistDataStoreOutput
End Property

Friend Property Get IsLiveOrderSubmitterConfigured() As Boolean
Const ProcName As String = "IsLiveOrderSubmitterConfigured"
On Error GoTo Err

IsLiveOrderSubmitterConfigured = (Not GetServiceProviderByRole(SPRoleOrderSubmissionLive) Is Nothing)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get IsAnyServiceProviderConfigured() As Boolean
IsAnyServiceProviderConfigured = mMaxSPIndex >= 0
End Property

Friend Property Get MarketDataManager() As IMarketDataManager
Set MarketDataManager = mMarketDataManager
End Property

Friend Property Get OrderPersistenceDataStore() As IOrderPersistenceDataStore
Set OrderPersistenceDataStore = mOrderPersistenceDataStore
End Property

Friend Property Get OrderRecoveryAgent() As IOrderRecoveryAgent
Set OrderRecoveryAgent = mOrderRecoveryAgent
End Property

Friend Property Get OrderSubmitterFactoryLive() As IOrderSubmitterFactory
Set OrderSubmitterFactoryLive = mOrderSubmitterFactoryLive
End Property

Friend Property Get OrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Set OrderSubmitterFactorySimulated = mOrderSubmitterFactorySimulated
End Property

Friend Property Let PermittedServiceProviderRoles(ByVal Value As ServiceProviderRoles)
mPermittedServiceProviderRoles = Value
End Property

Friend Property Get PermittedServiceProviderRoles() As ServiceProviderRoles
PermittedServiceProviderRoles = mPermittedServiceProviderRoles
End Property

Friend Property Get SupportedInputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedInputTickfileFormats"
On Error GoTo Err

SupportedInputTickfileFormats = GetSupportedTickfileFormats(SPRoleTickfileInput, TickfileStoreCapabilities.TickfileStoreCanReplay)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get SupportedOutputTickfileFormats() As TickfileFormatSpecifier()
Const ProcName As String = "SupportedOutputTickfileFormats"
On Error GoTo Err

SupportedOutputTickfileFormats = GetSupportedTickfileFormats(SPRoleTickfileOutput, TickfileStoreCapabilities.TickfileStoreCanRecord)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TickfileStoreInput() As ITickfileStore
Set TickfileStoreInput = mTickfileStoreInput
End Property

Friend Property Get TickfileStoreOutput() As ITickfileStore
Set TickfileStoreOutput = mTickfileStoreOutput
End Property

Friend Property Let TradeBuildAPI(ByVal Value As TradeBuildAPI)
Set mTradeBuildAPIRef = CreateWeakReference(Value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal progId As String, _
                ByVal Enabled As Boolean, _
                ByVal paramString As String, _
                Optional ByVal Name As String, _
                Optional ByVal Description As String) As IServiceProvider
Const ProcName As String = "Add"
On Error GoTo Err

Dim Params As Parameters
Dim spIndex As Long

Assert Not mServiceProvidersStarted, "Service providers have already been set up"

Set Add = CreateObject(progId)
Set Params = CreateParametersFromString(Replace(paramString, "\", "\\"))

If Enabled Then spIndex = AddSP(Add, Params, Name)

Set mServiceProviders(spIndex).SpConfig = AddEntryToConfig(progId, _
                                                        Enabled, _
                                                        Params, _
                                                        Name, _
                                                        Description)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ConfigureDefaultServiceProviders( _
                ByVal pPermittedServiceProviderRoles As ServiceProviderRoles, _
                ByVal pTwsServer As String, _
                ByVal pTwsPort As Long, _
                ByVal pTwsClientId As Long, _
                ByVal pTwsConnectionRetryIntervalSecs As Long, _
                ByVal pTwsLogLevel As String, _
                ByVal pTickfilesPath As String, _
                ByVal pOrderPersistenceFilePath As String)
Const ProcName As String = "ConfigureDefaultServiceProviders"
On Error GoTo Err

Assert Not IsAnyServiceProviderConfigured, "Service providers have already been set up"

mPermittedServiceProviderRoles = pPermittedServiceProviderRoles

If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleRealtimeData Then
    Add "IBTWSSP27.RealtimeDataServiceProvider", _
        True, _
        "Server=" & pTwsServer & _
        ";Port=" & pTwsPort & _
        ";Client Id=" & pTwsClientId & _
        ";Connection Retry Interval Secs=" & pTwsConnectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & pTwsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataPrimary Then
    Add "IBTWSSP27.ContractInfoServiceProvider", _
        True, _
        "Server=" & pTwsServer & _
        ";Port=" & pTwsPort & _
        ";Client Id=" & pTwsClientId & _
        ";Connection Retry Interval Secs=" & pTwsConnectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & pTwsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleContractDataSecondary Then
    ' not relevant in default scenario
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataInput Then
    Add "IBTWSSP27.HistDataServiceProvider", _
        True, _
        "Server=" & pTwsServer & _
        ";Port=" & pTwsPort & _
        ";Client Id=" & pTwsClientId & _
        ";Connection Retry Interval Secs=" & pTwsConnectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & pTwsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleHistoricalDataOutput Then
    ' no sensible default for this
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderPersistence Then
    Add ProgIdTbOrderPersistence, _
        True, _
        PropertyNameOrderPersistenceFilePath & pOrderPersistenceFilePath
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionLive Then
    Add "IBTWSSP27.OrderSubmissionSrvcProvider", _
        True, _
        "Server=" & pTwsServer & _
        ";Port=" & pTwsPort & _
        ";Client Id=" & pTwsClientId & _
        ";Connection Retry Interval Secs=" & pTwsConnectionRetryIntervalSecs & _
        ";Keep Connection=True" & _
        ";TWS Log Level=" & pTwsLogLevel
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleOrderSubmissionSimulated Then
    Dim lOrderSimulatorSP As New OrderSimulatorSP
    lOrderSimulatorSP.Initialise Me.MarketDataManager, Me.ContractStorePrimary, Me.ContractStoreSecondary
    AddSP lOrderSimulatorSP, Nothing, ""

End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileInput Then
    Add "TickfileSP27.TickfileServiceProvider", _
        True, _
        "Access Mode=Read only" & _
        ";Tickfile Path=" & pTickfilesPath
End If
If mPermittedServiceProviderRoles And ServiceProviderRoles.SPRoleTickfileOutput Then
    Add "TickfileSP27.TickfileServiceProvider", _
        True, _
        "Access Mode=Write only" & _
        ";Tickfile Path=" & pTickfilesPath
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetServiceProviderByRole( _
                ByVal pRole As ServiceProviderRoles, _
                Optional ByVal pInstance As Long = 1) As Object
Const ProcName As String = "GetServiceProviderByRole"
On Error GoTo Err

Dim i As Long
Dim lInstance As Long

AssertArgument pInstance > 0, "pInstance must be > 0"

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPRoles And pRole) = pRole Then
            lInstance = lInstance + 1
            If lInstance = pInstance Then
                Set GetServiceProviderByRole = mServiceProviders(i).SPObject
                Exit For
            End If
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub LoadServiceProvidersFromConfiguration()
Const ProcName As String = "LoadServiceProvidersFromConfiguration"
On Error GoTo Err

Assert Not mConfig Is Nothing, "No configuration has been supplied"

Set mServiceProviderConfig = mConfig.GetConfigurationSection(ConfigSectionServiceProviders)

If noServiceProvidersDefinedInConfig(mServiceProviderConfig) Then
    logMessage "No service providers defined in this configuration"
Else
    loadConfiguration
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function NameFromHandle( _
                ByVal Handle As Long) As String
Const ProcName As String = "NameFromHandle"
On Error GoTo Err

If Handle < 0 Then Exit Function
NameFromHandle = mServiceProviders(Handle).Name

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Remove( _
                ByVal sp As IServiceProvider)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim i As Long

If sp Is Nothing Then Exit Sub

For i = 0 To mMaxSPIndex
    If mServiceProviders(i).SPObject Is sp Then
        mServiceProviders(i).SPObject.Terminate
        RemoveSP i
        Exit For
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAll()
Const ProcName As String = "RemoveAll"
On Error GoTo Err

Dim i As Long
Dim sp As IServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        Set sp = mServiceProviders(i).SPObject
        sp.Terminate
        RemoveSP i
    End If
Next

mMaxSPIndex = -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener( _
                ByVal pListener As CollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function StartServiceProviders(ByVal pStudyLibraryManager As StudyLibraryManager) As Boolean
Const ProcName As String = "StartServiceProviders"
On Error GoTo Err

Assert Not mServiceProvidersStarted, "Service providers have already been started"

gLogger.Log "Starting service providers", ProcName, ModuleName

Dim lErrorOccurred As Boolean

Dim i As Long
For i = 0 To mMaxSPIndex
    gLogger.Log "Starting: " & mServiceProviders(i).Name, ProcName, ModuleName
    Dim lServiceProvider As IServiceProvider
    Set lServiceProvider = mServiceProviders(i).SPObject
    
    On Error Resume Next
    lServiceProvider.Link Me, i, mServiceProviders(i).Params
    If Err.Number <> 0 Then lErrorOccurred = True
    Err.Clear
    On Error GoTo Err
    
    fireCollectionChange CollItemAdded, lServiceProvider
Next

If lErrorOccurred Then
    gLogger.Log "One or more service providers failed to start", ProcName, ModuleName, LogLevelSevere
    StartServiceProviders = False
Else
    Set mContractStorePrimary = createContractStorePrimary
    Set mContractStoreSecondary = createContractStoreSecondary
    Set mHistDataStoreInput = createHistoricalDataStoreInput
    Set mHistDataStoreOutput = createHistoricalDataStoreOutput
    Set mOrderPersistenceDataStore = createOrderPersistenceDataStore
    Set mOrderRecoveryAgent = createOrderRecoveryAgent
    Set mOrderSubmitterFactoryLive = createOrderSubmitterFactoryLive
    Set mOrderSubmitterFactorySimulated = createOrderSubmitterFactorySimulated
    Set mMarketDataManager = createMarketDataManager(pStudyLibraryManager)
    Set mTickfileStoreInput = createTickfileStoreInput
    Set mTickfileStoreOutput = createTickfileStoreOutput
    
    gLogger.Log "Service providers started successfully", ProcName, ModuleName
    StartServiceProviders = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddEntryToConfig( _
            ByVal progId As String, _
            ByVal Enabled As Boolean, _
            ByVal Params As Parameters, _
            ByVal Name As String, _
            ByVal Description As String) As ConfigurationSection
Const ProcName As String = "AddEntryToConfig"
On Error GoTo Err

Dim sp As ConfigurationSection
Dim props As ConfigurationSection
Dim prop As ConfigurationSection
Dim param As Parameter

If mConfig Is Nothing Then Exit Function

If Name = "" Then Name = GenerateGUIDString

If mServiceProviderConfig Is Nothing Then
    Set mServiceProviderConfig = mConfig.AddConfigurationSection(ConfigSectionServiceProviders, ServiceProvidersRenderer)
End If

Set sp = mServiceProviderConfig.AddConfigurationSection(ConfigSectionServiceProvider & _
                                                            "(" & Name & ")", _
                                                            Description)
sp.setAttribute AttributeNameServiceProviderEnabled, CStr(Enabled)
sp.setAttribute AttributeNameServiceProviderProgId, progId

Set props = sp.AddConfigurationSection(ConfigSectionProperties)

For Each param In Params
    Set prop = props.AddConfigurationSection(ConfigSectonProperty & _
                                            "(" & param.Name & ")")
    prop.Value = param.Value
Next

Set AddEntryToConfig = sp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function AddSP( _
            ByVal serviceProvider As IServiceProvider, _
            ByVal Params As Parameters, _
            ByVal Name As String) As Long
Const ProcName As String = "AddSP"
On Error GoTo Err

Dim i As Long

If Name <> "" Then serviceProvider.Name = Name

' we don't reuse Removed entries to ensure that service providers Added later are
' always later in the table than those Added earlier
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        AssertArgument Not mServiceProviders(i).SPObject Is serviceProvider, "This Service Provider object has already been Added"
        AssertArgument mServiceProviders(i).Name <> serviceProvider.Name, "A Service Provider with this Name has already been Added"
    End If
Next

If mMaxSPIndex = UBound(mServiceProviders) Then
    ReDim Preserve mServiceProviders(UBound(mServiceProviders) + 10) As ServiceProviderTableEntry
End If
mMaxSPIndex = mMaxSPIndex + 1

mServiceProviders(mMaxSPIndex).Name = serviceProvider.Name
Set mServiceProviders(mMaxSPIndex).Params = Params
Set mServiceProviders(mMaxSPIndex).SPObject = serviceProvider

AddSP = mMaxSPIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub checkSPHasPermittedRole(ByVal pRole As ServiceProviderRoles, ByVal pHandle As Long)
Const ProcName As String = "checkSPHasPermittedRole"
On Error GoTo Err

If pRole And mPermittedServiceProviderRoles Then Exit Sub
Dim s As String
s = "Service provider " & NameFromHandle(pHandle) & " has impermissible role"
gLogger.Log s, ProcName, ModuleName
Assert True, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createContractStorePrimary() As IContractStore
Const ProcName As String = "createContractStorePrimary"
On Error GoTo Err

Dim lSp As IContractSP
Set lSp = GetServiceProviderByRole(SPRoleContractDataPrimary)

If lSp Is Nothing Then Exit Function

Set createContractStorePrimary = lSp.GetContractStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createContractStoreSecondary() As IContractStore
Const ProcName As String = "createContractStoreSecondary"
On Error GoTo Err

Dim lSp As IContractSP

Set lSp = GetServiceProviderByRole(SPRoleContractDataSecondary)
If lSp Is Nothing Then Exit Function

Set createContractStoreSecondary = lSp.GetContractStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createHistoricalDataStoreInput() As IHistoricalDataStore
Const ProcName As String = "createHistoricalDataStore"
On Error GoTo Err

Dim lSp As IHistDataSP
Set lSp = GetServiceProviderByRole(SPRoleHistoricalDataInput)

If lSp Is Nothing Then Exit Function

Set createHistoricalDataStoreInput = lSp.GetHistoricalDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                            
Private Function createHistoricalDataStoreOutput() As IHistoricalDataStore
Const ProcName As String = "createHistoricalDataStoreOutput"
On Error GoTo Err

Dim lSp As IHistDataSP
Set lSp = GetServiceProviderByRole(SPRoleHistoricalDataOutput)

If lSp Is Nothing Then Exit Function

Set createHistoricalDataStoreOutput = lSp.GetHistoricalDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                            
Private Function createMarketDataManager(ByVal pStudyLibraryManager As StudyLibraryManager) As IMarketDataManager
Const ProcName As String = "createMarketDataManager"
On Error GoTo Err

Dim lMarketDataFactory As IMarketDataFactory
Set lMarketDataFactory = GetMarketDataFactory
If lMarketDataFactory Is Nothing Then Exit Function

Set createMarketDataManager = CreateRealtimeDataManager(GetMarketDataFactory, pStudyLibraryManager, tb.MarketDataSourceOptions, tb.DefaultStateChangeListener, tb.NumberOfMarketDepthRows)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createOrderPersistenceDataStore() As IOrderPersistenceDataStore
Const ProcName As String = "createOrderPersistenceDataStore"
On Error GoTo Err

Dim lSp As IOrderPersistenceSP

Set lSp = GetServiceProviderByRole(SPRoleOrderPersistence)
If lSp Is Nothing Then Exit Function

Set createOrderPersistenceDataStore = lSp.GetOrderPersistenceDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createOrderRecoveryAgent() As IOrderRecoveryAgent
Const ProcName As String = "createOrderRecoveryAgent"
On Error GoTo Err

Dim lSp As IOrderSubmissionSP

Set lSp = GetServiceProviderByRole(SPRoleOrderSubmissionLive)
If lSp Is Nothing Then Exit Function

Set createOrderRecoveryAgent = lSp.GetOrderRecoveryAgent

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createOrderSubmitterFactoryLive() As IOrderSubmitterFactory
Const ProcName As String = "createOrderSubmitterFactoryLive"
On Error GoTo Err

Dim lSp As IOrderSubmissionSP

Set lSp = GetServiceProviderByRole(SPRoleOrderSubmissionLive)
If lSp Is Nothing Then Exit Function

Set createOrderSubmitterFactoryLive = lSp.GetOrderSubmitterFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createOrderSubmitterFactorySimulated() As IOrderSubmitterFactory
Const ProcName As String = "createOrderSubmitterFactorySimulated"
On Error GoTo Err

Dim lSp As IOrderSubmissionSP

Set lSp = GetServiceProviderByRole(SPRoleOrderSubmissionSimulated)
If lSp Is Nothing Then Exit Function

Set createOrderSubmitterFactorySimulated = lSp.GetOrderSubmitterFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createTickfileStoreInput() As ITickfileStore
Const ProcName As String = "createTickfileStoreInput"
On Error GoTo Err

Dim lSp As ITickfileSP
Set lSp = GetServiceProviderByRole(SPRoleTickfileInput)

If lSp Is Nothing Then Exit Function

Set createTickfileStoreInput = lSp.GetTickfileStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                            
Private Function createTickfileStoreOutput() As ITickfileStore
Const ProcName As String = "createTickfileStoreOutput"
On Error GoTo Err

Dim lSp As ITickfileSP
Set lSp = GetServiceProviderByRole(SPRoleTickfileOutput)

If lSp Is Nothing Then Exit Function

Set createTickfileStoreOutput = lSp.GetTickfileStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                            
Private Sub fireCollectionChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As IServiceProvider)
Const ProcName As String = "fireCollectionChange"
On Error GoTo Err

Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEventData

Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject
For Each listener In mCollectionChangeListeners.CurrentListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function GetMarketDataFactory() As IMarketDataFactory
Const ProcName As String = "GetMarketDataFactory"
On Error GoTo Err

Dim lSp As IRealtimeDataSP

Set lSp = GetServiceProviderByRole(SPRoleRealtimeData)
If lSp Is Nothing Then Exit Function

Set GetMarketDataFactory = lSp.GetMarketDataFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function GetSupportedTickfileFormats( _
                ByVal pRole As ServiceProviderRoles, _
                ByVal pCapabilities As TickfileStoreCapabilities) As TickfileFormatSpecifier()
Const ProcName As String = "GetSupportedTickfileFormats"
On Error GoTo Err

Dim i As Long
Dim j As Long
Dim k As Long
Dim lTickfileSp As ITickfileSP
Dim lSupportedFormats() As TickfileFormatSpecifier

ReDim lAllSupportedFormats(9) As TickfileFormatSpecifier

k = -1
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPRoles And pRole) <> 0 Then
            Set lTickfileSp = mServiceProviders(i).SPObject
            lSupportedFormats = lTickfileSp.GetTickfileStore.SupportedFormats
            For j = 0 To UBound(lSupportedFormats)
                If (lSupportedFormats(j).Capabilities And pCapabilities) <> 0 Then
                    k = k + 1
                    If k > UBound(lAllSupportedFormats) Then
                        ReDim Preserve lAllSupportedFormats(UBound(lAllSupportedFormats) + 10) As TickfileFormatSpecifier
                    End If
                    lAllSupportedFormats(k) = lSupportedFormats(j)
                End If
            Next
        End If
    End If
Next

If k >= 0 Then
    ReDim Preserve lAllSupportedFormats(k) As TickfileFormatSpecifier
    GetSupportedTickfileFormats = lAllSupportedFormats
Else
    Erase lAllSupportedFormats
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSPWithUnexpectedType(ByVal pExpectedTypeName As String, ByVal pHandle As Long)
Const ProcName As String = "handleSPWithUnexpectedType"
On Error GoTo Err

Dim s As String
s = "Service provider " & NameFromHandle(pHandle) & " is not of type " & pExpectedTypeName
gLogger.Log s, ProcName, ModuleName, LogLevelSevere
Assert True, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadConfiguration()
Const ProcName As String = "loadConfiguration"
On Error GoTo Err

Dim sp As ConfigurationSection
Dim spName As String
Dim spProgId As String
Dim props As ConfigurationSection
Dim prop As ConfigurationSection
Dim spObj As Object
Dim spIndex As Long
Dim Params As Parameters

For Each sp In mServiceProviderConfig
    If CBool(sp.getAttribute(AttributeNameEnabled, "False")) Then
        
        spName = sp.InstanceQualifier
        
        spProgId = sp.getAttribute(AttributeNameServiceProviderProgId)
        AssertArgument spProgId <> "", "Service provider ProgId not specified"
        
        Set Params = New Parameters
        
        On Error Resume Next
        Set props = Nothing
        Set props = sp.GetConfigurationSection(ConfigSectionProperties)
        On Error GoTo Err
        
        If Not props Is Nothing Then
            For Each prop In props
                AssertArgument prop.InstanceQualifier <> "", "Property Name not specified"
                Params.SetParameterValue prop.InstanceQualifier, prop.Value
            Next
        End If
        
        On Error Resume Next
        Set spObj = Nothing
        If spProgId = ProgIdTbOrders Then
            ' this is the built-in order simulator
            Set spObj = New OrderSimulatorSP
        Else
            Set spObj = CreateObject(spProgId)
        End If
        On Error GoTo Err
        AssertArgument Not spObj Is Nothing, "Can't create service provider with progId " & spProgId
        
        spIndex = AddSP(spObj, Params, spName)
        AssertArgument spIndex >= 0, "Can't Add service provider " & spName
            
        Set mServiceProviders(spIndex).SpConfig = sp
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function noServiceProvidersDefinedInConfig(ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "noServiceProvidersDefinedInConfig"
On Error GoTo Err

If pConfig Is Nothing Then
    noServiceProvidersDefinedInConfig = True
    Exit Function
End If

Dim en As Enumerator
Set en = pConfig.ChildEnumerator

noServiceProvidersDefinedInConfig = Not en.MoveNext

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub RemoveSP( _
                ByVal Handle As Long, _
                Optional ByVal RemoveFromConfig As Boolean)
Const ProcName As String = "RemoveSP"
On Error GoTo Err

Dim sp As IServiceProvider

If Handle < 0 Then Exit Sub

Set sp = mServiceProviders(Handle).SPObject

If RemoveFromConfig And Not mServiceProviders(Handle).SpConfig Is Nothing Then
    mServiceProviderConfig.RemoveConfigurationSection mServiceProviders(Handle).SpConfig
End If

Set mServiceProviders(Handle).SPObject = Nothing
Set mServiceProviders(Handle).SpConfig = Nothing
mServiceProviders(Handle).SPRoles = 0
mServiceProviders(Handle).Name = ""
fireCollectionChange CollItemRemoved, sp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function tb() As TradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function

