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
Implements TradeBuildSP.IContractInfoServiceConsumer

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName As String = "TradeBuild26"
Private Const ModuleName As String = "ServiceProviders"

Private Const AttributeNameEnabled As String = "Enabled"
Private Const AttributeNameLogLevel As String = "LogLevel"
Private Const AttributeNamePropertyName As String = "Name"
Private Const AttributeNamePropertyValue As String = "Value"
Private Const AttributeNameServiceProviderEnabled As String = "Enabled"
Private Const AttributeNameServiceProviderName As String = "Name"
Private Const AttributeNameServiceProviderProgId As String = "ProgId"

Private Const ConfigNameProperties As String = "Properties"
Private Const ConfigNameProperty As String = "Property"
Private Const ConfigNameServiceProvider As String = "ServiceProvider"
Private Const ConfigNameServiceProviders As String = "ServiceProviders"

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
End Type
'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviders()         As ServiceProviderTableEntry
Private mMaxSPIndex                 As Long

Private mChangeListeners            As Collection

Private mConfig                     As ConfigItem
Private mServiceProviderConfigList  As ConfigItem

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mServiceProviders(9) As ServiceProviderTableEntry
mMaxSPIndex = -1

Set mChangeListeners = New Collection

End Sub

'@================================================================================
' ICommonServiceConsumer Interface Members
'@================================================================================

Private Sub ICommonServiceConsumer_FatalServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
gTB.NotifyFatalError errorCode, _
                    errorMessage, _
                    nameFromHandle(spHandle) & ":" & source
removeSP spHandle
End Sub

Private Sub ICommonServiceConsumer_NotifyListeners( _
                ByVal value As Variant, _
                ByVal valueType As Long, _
                ByVal source As Object)
gTB.notifyListeners value, valueType, source
End Sub

Private Sub ICommonServiceConsumer_RegisterServiceProvider( _
                            ByVal handle As Long, _
                            ByVal spType As TradeBuildSP.ServiceProviderTypes)
Dim regOk As Boolean
mServiceProviders(handle).SPTypes = mServiceProviders(handle).SPTypes Or spType
Select Case spType
Case TradeBuildSP.ServiceProviderTypes.ContractInfo
    Dim contractInfoSP As TradeBuildSP.IContractInfoServiceProvider
    
    On Error Resume Next
    Set contractInfoSP = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not contractInfoSP Is Nothing Then
        contractInfoSP.ContractInfoServiceConsumer = Me
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.ExchangeSimulator
    Dim exchangeSimulatorSp As TradeBuildSP.IOrderSubmissionSrvcProvider
    
    On Error Resume Next
    Set exchangeSimulatorSp = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not exchangeSimulatorSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.HistoricalData
    Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
    
    On Error Resume Next
    Set histDataSp = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not histDataSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.LiveOrderSubmission
    Dim liveOrderSp As TradeBuildSP.IOrderSubmissionSrvcProvider
    
    On Error Resume Next
    Set liveOrderSp = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not liveOrderSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.OrderPersistence
Case TradeBuildSP.ServiceProviderTypes.RealtimeData
    Dim realtimeDataSp As TradeBuildSP.IRealtimeDataServiceProvider
    
    On Error Resume Next
    Set realtimeDataSp = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not realtimeDataSp Is Nothing Then
        regOk = True
    End If
Case TradeBuildSP.ServiceProviderTypes.Tickfile
    Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
    
    On Error Resume Next
    Set tickfileSp = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not tickfileSp Is Nothing Then regOk = True
End Select

If Not regOk Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild26.ServiceProviders::ICommonServiceConsumer_RegisterServiceProvider", _
                "Caller is not a type " & spType & " service provider object"
End If
    
mServiceProviders(handle).SPTypes = mServiceProviders(handle).SPTypes Or spType
End Sub

Private Sub ICommonServiceConsumer_ServiceProviderError( _
                            ByVal errorCode As Long, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
gTB.SetServiceProviderError errorCode, _
                errorMessage, _
                nameFromHandle(spHandle)
End Sub

'@================================================================================
' IContractInfoServiceConsumer Interface Members
'@================================================================================

Private Sub IContractInfoServiceConsumer_LoadComplete( _
                            ByVal handle As Long, _
                            ByVal pContractsbuilder As ContractsBuilder)
Dim lContracts As Contracts

Set lContracts = pContractsbuilder.Contracts
If lContracts.Count = 0 Then
    ' try another service provider if any
    If Not LoadContractsFromNextSP(handle + 1, pContractsbuilder) Then
        pContractsbuilder.LoadComplete
    End If
Else
    pContractsbuilder.LoadComplete
End If
End Sub

Private Sub IContractInfoServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal pContractsbuilder As ContractsBuilder, _
                ByVal serviceProviderHandle As Long)

Dim ev As NotificationEvent

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.CIConnectedToDataSource
    ev.eventCode = eventCode
    ev.eventMessage = eventMessage
    Set ev.source = pContractsbuilder.Contracts
    pContractsbuilder.notify ev
Case TradeBuildSP.StandardSPEventCodes.CIContractSpecifierInvalid
    pContractsbuilder.ContractSpecifierInvalid eventMessage
Case TradeBuildSP.StandardSPEventCodes.CICantConnectDataSource
    gTB.SetServiceProviderError eventCode, _
                        "Can't connect to Contract database: " & eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIDisconnectedFromDataSource
    gTB.SetServiceProviderError eventCode, _
                        "Disconnected from Contract database: " & eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CILostConnectionToDataSource
    gTB.SetServiceProviderError eventCode, _
                        "Lost connection to Contract database: " & eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIReConnectingDataSource
    gTB.SetServiceProviderError eventCode, _
                        "Reconnecting to Contract database: " & eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIRetryConnectDataSource
    gTB.SetServiceProviderError eventCode, _
                        "Attempting again to connect to Contract database: " & eventMessage, _
                        gTB.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case Else
    gTB.SetServiceProviderError eventCode, _
                        eventMessage, _
                        nameFromHandle(serviceProviderHandle)
End Select
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

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

Public Function add( _
            ByVal progId As String, _
            ByVal enabled As Boolean, _
            ByVal paramString As String, _
            Optional ByVal logLevel As LogLevels = LogLevelLow, _
            Optional ByVal name As String, _
            Optional ByVal Description As String _
            ) As TradeBuildSP.ICommonServiceProvider
Dim params As parameters
Set add = CreateObject(progId)
Set params = CreateParametersFromString(paramString)

If enabled Then
    If Not addSP(add, params, logLevel, name) Then
        Set add = Nothing
        Exit Function
    End If
End If

addEntryToConfig progId, _
                enabled, _
                params, _
                logLevel, _
                name, _
                Description
End Function

Public Sub addCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
mChangeListeners.add value
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
                ByVal contractSpec As contractSpecifier, _
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
                                                    contractSpec, _
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
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.LiveOrderSubmission) <> 0 Then
            Set lOrderSimulatorSp = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

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
    ' no order simulator sp has been added, so we will resort
    ' to using the default built-in simulator
    Set lOrderSimulatorSp = New OrderSimulatorSP
    addSP lOrderSimulatorSp, Nothing, LogLevelLow, ""
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
                                                    outputTickfileFormatID, _
                                                    outputTickfilePath)

tfWriter.ServiceProviderWriter = tickfileSPWriter

Set CreateTickfileWriter = tfWriter

End Function

Friend Function IsSupportedHistoricalDataPeriod(ByRef period As TimePeriod) As Boolean
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

Friend Sub loadContracts( _
                ByVal pContractsbuilder As ContractsBuilder)

If Not LoadContractsFromNextSP(0, pContractsbuilder) Then pContractsbuilder.LoadComplete
End Sub

Public Sub loadServiceProviderConfiguration( _
                value As ConfigItem)
Set mConfig = value

On Error Resume Next
Set mServiceProviderConfigList = mConfig.childItems.item(ConfigNameServiceProviders)
On Error GoTo 0

If mServiceProviderConfigList Is Nothing Then
    Set mServiceProviderConfigList = mConfig.childItems.AddItem(ConfigNameServiceProviders)
End If

loadConfiguration
End Sub

Friend Function nameFromHandle( _
                ByVal handle As Long) As String
If handle < 0 Then Exit Function
nameFromHandle = mServiceProviders(handle).name
End Function

Public Sub remove( _
                ByVal sp As TradeBuildSP.ICommonServiceProvider)
Dim i As Long

If sp Is Nothing Then Exit Sub

For i = 0 To mMaxSPIndex
    If mServiceProviders(i).SPObject Is sp Then
        mServiceProviders(i).SPObject.Terminate
        removeSP i
        Exit For
    End If
Next
End Sub

Public Sub RemoveAll()
Dim i As Long
Dim sp As TradeBuildSP.ICommonServiceProvider

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        Set sp = mServiceProviders(i).SPObject
        sp.Terminate
        Set mServiceProviders(i).SPObject = Nothing
        mServiceProviders(i).SPTypes = 0
        mServiceProviders(i).name = ""
        fireChange CollItemRemoved, sp
    End If
Next
End Sub

Public Sub removeCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.item(i) Is value Then mChangeListeners.remove i
Next
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addEntryToConfig( _
            ByVal progId As String, _
            ByVal enabled As Boolean, _
            ByVal params As parameters, _
            ByVal logLevel As LogLevels, _
            ByVal name As String, _
            ByVal Description As String _
            )
Dim sp As ConfigItem
Dim props As ConfigItem
Dim prop As ConfigItem
Dim param As Parameter

If mConfig Is Nothing Then Exit Sub

Set sp = mServiceProviderConfigList.childItems.AddItem(ConfigNameServiceProvider, _
                                                        , _
                                                        Description)
sp.setAttribute AttributeNameServiceProviderEnabled, CStr(enabled)
If name <> "" Then sp.setAttribute AttributeNameServiceProviderName, name
sp.setAttribute AttributeNameServiceProviderProgId, progId
sp.setAttribute AttributeNameLogLevel, logLevel

Set props = sp.childItems.AddItem(ConfigNameProperties)

For Each param In params
    Set prop = props.childItems.AddItem(ConfigNameProperty)
    prop.setAttribute AttributeNamePropertyName, param.name
    prop.setAttribute AttributeNamePropertyValue, param.value
Next

End Sub

Private Function addSP( _
            ByVal serviceProvider As TradeBuildSP.ICommonServiceProvider, _
            ByVal params As parameters, _
            ByVal logLevel As LogLevels, _
            ByVal name As String _
            ) As Boolean
Dim i As Long

If name <> "" Then serviceProvider.name = name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If mServiceProviders(i).SPObject Is serviceProvider Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.ServiceProviders::Add", _
                    "This Service Provider object has already been added"
        End If
        If mServiceProviders(i).name = serviceProvider.name Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild26.ServiceProviders::Add", _
                    "A Service Provider with this name has already been added"
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
    addSP = False
Else
    serviceProvider.logLevel = logLevel
    fireChange CollItemAdded, serviceProvider
    addSP = True
End If
End Function

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As TradeBuildSP.ICommonServiceProvider)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Set ev.source = Me
ev.changeType = changeType
Set ev.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)
End Sub

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
Dim sp As ConfigItem
Dim spName As String
Dim spProgId As String
Dim spLogLevel As String
Dim props As ConfigItem
Dim prop As ConfigItem
Dim propName As String
Dim propValue As String
Dim spObj As Object
Dim params As parameters

For Each sp In mServiceProviderConfigList.childItems
    If CBool(sp.getAttribute(AttributeNameEnabled)) Then
        
        On Error Resume Next

        spName = sp.getAttribute(AttributeNameServiceProviderName)
        
        spProgId = sp.getAttribute(AttributeNameServiceProviderProgId)
        If spProgId = "" Then
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & "loadConfiguration", _
                    "Service provider ProgId not specified in configuration file"
        End If
        
        spLogLevel = sp.getAttribute(AttributeNameLogLevel)
        If spLogLevel = "" Then spLogLevel = CStr(LogLevels.LogLevelLow)
        
        Set params = New parameters
        
        Set props = sp.childItems.item(ConfigNameProperties)
        
        If Not props Is Nothing Then
            For Each prop In props.childItems
                propName = prop.getAttribute(AttributeNamePropertyName)
                propValue = prop.getAttribute(AttributeNamePropertyValue)
                params.setParameterValue propName, propValue
            Next
        End If
        
        On Error GoTo 0
        
        Set spObj = CreateObject(spProgId)
        addSP spObj, params, spLogLevel, spName
    End If
Next
End Sub

Private Function LoadContractsFromNextSP( _
                            ByVal startingIndex As Long, _
                            ByVal pContractsbuilder As ContractsBuilder _
                            ) As Boolean
Dim i As Long
Dim contractInfoSP As TradeBuildSP.IContractInfoServiceProvider

For i = startingIndex To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.ContractInfo) <> 0 Then
            Set contractInfoSP = mServiceProviders(i).SPObject
            Exit For
        End If
    End If
Next

If contractInfoSP Is Nothing Then Exit Function

contractInfoSP.RequestContractDetails pContractsbuilder
LoadContractsFromNextSP = True
End Function

Private Sub removeSP( _
                ByVal handle As Long)
Dim sp As ICommonServiceProvider
Set sp = mServiceProviders(handle).SPObject
Set mServiceProviders(handle).SPObject = Nothing
mServiceProviders(handle).SPTypes = 0
mServiceProviders(handle).name = ""
fireChange CollItemRemoved, sp
End Sub
