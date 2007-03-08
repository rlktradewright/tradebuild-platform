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

Private mTradeBuildAPIRef           As WeakReference

Private mServiceProviders() As ServiceProviderTableEntry
Private mMaxSPIndex As Long

Private mChangeListeners As Collection

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
tb.notifyFatalError errorCode, _
                    errorMessage, _
                    nameFromHandle(spHandle) & ":" & source
End Sub

Private Function ICommonServiceConsumer_NewContract() As TradeBuildSP.IContract
Set ICommonServiceConsumer_NewContract = New Contract
End Function

Private Function ICommonServiceConsumer_NewContracts() As TradeBuildSP.IContracts
Set ICommonServiceConsumer_NewContracts = tb.NewContracts(Nothing)
End Function

Private Function ICommonServiceConsumer_NewContractSpecifier() As TradeBuildSP.IContractSpecifier
Set ICommonServiceConsumer_NewContractSpecifier = New contractSpecifier
End Function

Private Function ICommonServiceConsumer_NewExecution() As TradeBuildSP.IExecution
Set ICommonServiceConsumer_NewExecution = New Execution
End Function

Private Sub ICommonServiceConsumer_NotifyListeners( _
                ByVal value As Variant, _
                ByVal valueType As Long, _
                ByVal source As Object)
tb.notifyListeners value, valueType, source
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
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild25.ServiceProviders::ICommonServiceConsumer_RegisterServiceProvider", _
                "Caller is not a type " & spType & " service provider object"
End If
    
mServiceProviders(handle).SPTypes = mServiceProviders(handle).SPTypes Or spType
End Sub

Private Sub ICommonServiceConsumer_ServiceProviderError( _
                            ByVal errorCode As Long, _
                            ByVal errorMessage As String, _
                            ByVal spHandle As Long)
tb.SetServiceProviderError errorCode, _
                errorMessage, _
                nameFromHandle(spHandle)
End Sub

'@================================================================================
' IContractInfoServiceConsumer Interface Members
'@================================================================================

Private Sub IContractInfoServiceConsumer_LoadComplete( _
                            ByVal handle As Long, _
                            ByVal pContracts As TradeBuildSP.IContracts)
Dim lContracts As TradeBuild25.Contracts

Set lContracts = pContracts
If lContracts.Count = 0 Then
    ' try another service provider if any
    If Not LoadContractsFromNextSP(handle + 1, pContracts) Then
        lContracts.LoadComplete
    End If
Else
    lContracts.LoadComplete
End If
End Sub

Private Sub IContractInfoServiceConsumer_NotifyEvent( _
                ByVal eventCode As TradeBuildSP.StandardSPEventCodes, _
                ByVal eventMessage As String, _
                ByVal pContracts As TradeBuildSP.IContracts, _
                ByVal serviceProviderHandle As Long)
Dim lContracts As TradeBuild25.Contracts

Set lContracts = pContracts

Select Case eventCode
Case TradeBuildSP.StandardSPEventCodes.CIConnectedToDataSource
    Dim ev As NotificationEvent
    ev.eventCode = eventCode
    ev.eventMessage = eventMessage
    Set ev.source = lContracts
    lContracts.notify ev
Case TradeBuildSP.StandardSPEventCodes.CIContractSpecifierInvalid
    lContracts.ContractSpecifierInvalid eventMessage
Case TradeBuildSP.StandardSPEventCodes.CICantConnectDataSource
    tb.SetServiceProviderError eventMessage, _
                        "Can't connect to Contract database: " & eventMessage, _
                        tb.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIDisconnectedFromDataSource
    tb.SetServiceProviderError eventMessage, _
                        "Disconnected from Contract database: " & eventMessage, _
                        tb.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CILostConnectionToDataSource
    tb.SetServiceProviderError eventMessage, _
                        "Lost connection to Contract database: " & eventMessage, _
                        tb.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIReConnectingDataSource
    tb.SetServiceProviderError eventMessage, _
                        "Reconnecting to Contract database: " & eventMessage, _
                        tb.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case TradeBuildSP.StandardSPEventCodes.CIRetryConnectDataSource
    tb.SetServiceProviderError eventMessage, _
                        "Attempting again to connect to Contract database: " & eventMessage, _
                        tb.ServiceProviders.nameFromHandle(serviceProviderHandle)
Case Else
    tb.SetServiceProviderError eventMessage, _
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

Friend Property Get SupportedHistoricalDataPeriods() As timePeriod()
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

Friend Property Get SupportedInputTickfileFormats() As TradeBuild25.TickfileFormatSpecifier()
SupportedInputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Replay)
End Property

Friend Property Get SupportedOutputTickfileFormats() As TradeBuild25.TickfileFormatSpecifier()
SupportedOutputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Record)
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = weakrefs.createWeakReference(value)
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
            ByVal serviceProvider As TradeBuildSP.ICommonServiceProvider, _
            Optional ByVal logLevel As TradeBuild25.LogLevels, _
            Optional ByVal name As String _
            ) As TradeBuildSP.ICommonServiceProvider
Dim i As Long

If name <> "" Then serviceProvider.name = name

' we don't reuse removed entries to ensure that service providers added later are
' always later in the table than those added earlier
For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If mServiceProviders(i).SPObject Is serviceProvider Then
            err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild25.ServiceProviders::Add", _
                    "This Service Provider object has already been added"
        End If
        If mServiceProviders(i).name = serviceProvider.name Then
            err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild25.ServiceProviders::Add", _
                    "A Service Provider with this name has already been added"
        End If
    End If
Next

If mMaxSPIndex = UBound(mServiceProviders) Then
    ReDim Preserve mServiceProviders(UBound(mServiceProviders) + 10) As ServiceProviderTableEntry
End If
mMaxSPIndex = mMaxSPIndex + 1
Set add = serviceProvider

mServiceProviders(mMaxSPIndex).name = serviceProvider.name
Set mServiceProviders(mMaxSPIndex).SPObject = serviceProvider

serviceProvider.Link Me, mMaxSPIndex
serviceProvider.logLevel = logLevel

fireChange CollItemAdded, serviceProvider
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
                ByVal contractSpec As TradeBuild25.contractSpecifier, _
                ByVal historicDataPath As String _
                ) As TradeBuild25.HistoricDataWriter
                            
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
Dim histDataSPWriter As TradeBuildSP.IBarDataWriter
Dim histDataWriter As TradeBuild25.HistoricDataWriter
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

Set histDataWriter = New TradeBuild25.HistoricDataWriter
histDataWriter.ServiceProviderWriter = histDataSPWriter
Set CreateHistoricalDataWriter = histDataWriter

End Function

Friend Function CreateLiveOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As TradeBuild25.Contract _
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
                ByVal pContract As TradeBuild25.Contract) As TradeBuildSP.IOrderSubmitter
                            
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
    Set lOrderSimulatorSp = New TradeBuild25.OrderSimulatorSP
    add lOrderSimulatorSp
End If

Set orderSubmitter = lOrderSimulatorSp.CreateOrderSubmitter(serviceConsumer, pContract)

Set CreateSimulatedOrderSubmitter = orderSubmitter

                            
End Function

Friend Function CreateTickfileReader( _
                ByVal pTicker As Ticker, _
                ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                ByVal inputTickfileFormatID As String _
                ) As TradeBuild25.TickfileReader
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPReader As TradeBuildSP.ITickfileReader
Dim tfReader As TradeBuild25.TickfileReader
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

Set tfReader = New TradeBuild25.TickfileReader
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
            ByVal outputTickfilePath As String _
            ) As TradeBuild25.TickfileWriter
                            
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPWriter As TradeBuildSP.ITickfileWriter
Dim tfWriter As TradeBuild25.TickfileWriter
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

Set tfWriter = New TradeBuild25.TickfileWriter
Set tickfileSPWriter = tickfileSp.CreateTickfileWriter(serviceConsumer, _
                                                    outputTickfileFormatID, _
                                                    outputTickfilePath)

tfWriter.ServiceProviderWriter = tickfileSPWriter

Set CreateTickfileWriter = tfWriter

End Function

Friend Function IsSupportedHistoricalDataPeriod(ByRef period As timePeriod)
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

Friend Sub LoadContracts( _
                ByVal Contracts As TradeBuild25.Contracts)

If Not LoadContractsFromNextSP(0, Contracts) Then Contracts.LoadComplete
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
        Set mServiceProviders(i).SPObject = Nothing
        mServiceProviders(i).SPTypes = 0
        mServiceProviders(i).name = ""
        fireChange CollItemRemoved, sp
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
                            ) As TradeBuild25.TickfileFormatSpecifier()
Dim i As Long
Dim j As Long
Dim k As Long
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim supportedFormats() As TradeBuild25.TickfileFormatSpecifier
ReDim allSupportedFormats(9) As TradeBuild25.TickfileFormatSpecifier

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
                        ReDim Preserve allSupportedFormats(UBound(allSupportedFormats) + 10) As TradeBuild25.TickfileFormatSpecifier
                    End If
                    allSupportedFormats(k) = supportedFormats(j)
                End If
            Next
        End If
    End If
Next

If k >= 0 Then
    ReDim Preserve allSupportedFormats(k) As TradeBuild25.TickfileFormatSpecifier
    GetSupportedTickfileFormats = allSupportedFormats
Else
    Erase allSupportedFormats
End If
End Function


Private Function LoadContractsFromNextSP( _
                            ByVal startingIndex As Long, _
                            ByVal pContracts As TradeBuild25.Contracts _
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

contractInfoSP.RequestContractDetails pContracts
LoadContractsFromNextSP = True
End Function

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function

