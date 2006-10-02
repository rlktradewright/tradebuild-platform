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

Implements TradeBuildSP.ICommonServiceConsumer
Implements TradeBuildSP.IContractInfoServiceConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type ServiceProviderTableEntry
    name        As String
    SPTypes     As Long
    SPObject    As TradeBuildSP.ICommonServiceProvider
End Type
'================================================================================
' Member variables
'================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mServiceProviders() As ServiceProviderTableEntry
Private mMaxSPIndex As Long

Private mChangeListeners As WeakCollection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mServiceProviders(9) As ServiceProviderTableEntry
mMaxSPIndex = -1

Set mChangeListeners = New WeakCollection

End Sub

'================================================================================
' ICommonServiceConsumer Interface Members
'================================================================================

Private Sub ICommonServiceConsumer_FatalServiceProviderError( _
                ByVal errorCode As Long, _
                ByVal source As String, _
                ByVal errorMessage As String, _
                ByVal spHandle As Long)
tb.notifyFatalError errorCode, _
                    errorMessage, _
                    nameFromHandle(spHandle) & ":" & source
End Sub

Private Function ICommonServiceConsumer_FormatTimestamp( _
                ByVal timestamp As Date, _
                Optional ByVal formatOption As TradeBuildSP.TimestampFormats = 2&, _
                Optional ByVal formatString As String = "yyyymmddhhnnss" _
                ) As String
ICommonServiceConsumer_FormatTimestamp = gFormatTimestamp(timestamp, _
                                                    formatOption, _
                                                    formatString)
End Function

Private Function ICommonServiceConsumer_GenerateGUID() As TradeBuildSP.GUID
Dim lGuid As TradeBuild.GUIDStruct
lGuid = gGenerateGUID
ICommonServiceConsumer_GenerateGUID.data(0) = lGuid.data(0)
ICommonServiceConsumer_GenerateGUID.data(1) = lGuid.data(1)
ICommonServiceConsumer_GenerateGUID.data(2) = lGuid.data(2)
ICommonServiceConsumer_GenerateGUID.data(3) = lGuid.data(3)
End Function

Private Function ICommonServiceConsumer_GenerateGUIDString() As String
ICommonServiceConsumer_GenerateGUIDString = gGenerateGUIDString
End Function

Private Function ICommonServiceConsumer_GenerateId() As Long
ICommonServiceConsumer_GenerateId = gGenerateID
End Function

Private Function ICommonServiceConsumer_GenerateIdString() As String
ICommonServiceConsumer_GenerateIdString = gGenerateIDString
End Function

Private Function ICommonServiceConsumer_GetTimestamp() As Date
ICommonServiceConsumer_GetTimestamp = gCurrentTime
End Function

Private Function ICommonServiceConsumer_GUIDToString(pGUID As TradeBuildSP.GUID) As String
Dim lGuid As TradeBuild.GUIDStruct
lGuid.data(0) = pGUID.data(0)
lGuid.data(1) = pGUID.data(1)
lGuid.data(2) = pGUID.data(2)
lGuid.data(3) = pGUID.data(3)
ICommonServiceConsumer_GUIDToString = gGUIDToString(lGuid)
End Function

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

Private Function ICommonServiceConsumer_NewParameters() As TradeBuildSP.IParameters
Set ICommonServiceConsumer_NewParameters = New TradeBuild.parameters
End Function

Private Function ICommonServiceConsumer_NewStudies() As TradeBuildSP.iStudies
Set ICommonServiceConsumer_NewStudies = New studies
End Function

Private Function ICommonServiceConsumer_NewStudyDefinition() As TradeBuildSP.IStudyDefinition
Set ICommonServiceConsumer_NewStudyDefinition = New studyDefinition
End Function

Private Function ICommonServiceConsumer_NewStudyParameterDefinition() As TradeBuildSP.IStudyParameterDefinition
Set ICommonServiceConsumer_NewStudyParameterDefinition = New StudyParameterDefinition
End Function

Private Function ICommonServiceConsumer_NewStudyValueCache( _
                ByVal cyclicSize As Long, _
                ByVal ValueName As String) As TradeBuildSP.IStudyValueCache
Dim valueCache As New StudyValueCache
valueCache.cyclicSize = cyclicSize
valueCache.ValueName = ValueName
Set ICommonServiceConsumer_NewStudyValueCache = valueCache
End Function

Private Function ICommonServiceConsumer_NewStudyValueDefinition() As TradeBuildSP.IStudyValueDefinition
Set ICommonServiceConsumer_NewStudyValueDefinition = New StudyValueDefinition
End Function

Private Function ICommonServiceConsumer_NewStudyValueListeners() As TradeBuildSP.IStudyValueListeners
Set ICommonServiceConsumer_NewStudyValueListeners = New StudyValueListeners
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
Case TradeBuildSP.ServiceProviderTypes.studies
    Dim studySP As TradeBuildSP.IStudyServiceProvider
    
    On Error Resume Next
    Set studySP = mServiceProviders(handle).SPObject
    On Error GoTo 0
    If Not studySP Is Nothing Then regOk = True
End Select

If Not regOk Then
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuild.ServiceProviders::ICommonServiceConsumer_RegisterServiceProvider", _
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

Private Sub ICommonServiceConsumer_SortObjects(data() As TradeBuildSP.SortEntry)
Dim i As Long
ReDim ldata(LBound(data) To UBound(data)) As TradeBuild.SortEntryStruct

For i = LBound(data) To UBound(data)
    ldata(i).key = data(i).key
    ldata(i).ObjectRef = data(i).ObjectRef
Next

gSortObjects ldata, LBound(ldata), UBound(ldata)

For i = LBound(data) To UBound(data)
    data(i).key = ldata(i).key
    data(i).ObjectRef = ldata(i).ObjectRef
Next
End Sub

Private Function ICommonServiceConsumer_StartTask( _
                ByVal target As TradeBuildSP.ITask, _
                Optional ByVal name As String, _
                Optional ByVal data As Variant) As TradeBuildSP.ITaskCompletion
Dim Tasktarget As SpTask
Set Tasktarget = New SpTask
Tasktarget.SpTaskObject = target
Set ICommonServiceConsumer_StartTask = tb.taskManager.startTask(Tasktarget, name, data)
End Function

Private Function ICommonServiceConsumer_ToBytes(inString As String) As Byte()
ICommonServiceConsumer_ToBytes = gToBytes(inString)
End Function

Private Function ICommonServiceConsumer_ToHex(inAr() As Byte) As String
ICommonServiceConsumer_ToHex = gToHex(inAr)
End Function

'================================================================================
' IContractInfoServiceConsumer Interface Members
'================================================================================

Private Sub IContractInfoServiceConsumer_Error( _
                            ByVal errorCode As TradeBuildSP.StandardSPErrorCodes, _
                            ByVal errorMessage As String, _
                            ByVal pContracts As TradeBuildSP.IContracts, _
                            ByVal spHandle As Long)
Dim lContracts As TradeBuild.Contracts

Set lContracts = pContracts

Select Case errorCode
Case TradeBuildSP.StandardSPErrorCodes.CIContractSpecifierInvalid
    lContracts.ContractSpecifierInvalid errorMessage
Case TradeBuildSP.StandardSPErrorCodes.CICantConnectDataSource
    tb.SetServiceProviderError errorCode, _
                        "Can't connect to Contract database: " & errorMessage, _
                        tb.serviceProviders.nameFromHandle(spHandle)
Case TradeBuildSP.StandardSPErrorCodes.CIDisconnectedFromDataSource
    tb.SetServiceProviderError errorCode, _
                        "Disconnected from Contract database: " & errorMessage, _
                        tb.serviceProviders.nameFromHandle(spHandle)
Case TradeBuildSP.StandardSPErrorCodes.CILostConnectionToDataSource
    tb.SetServiceProviderError errorCode, _
                        "Lost connection to Contract database: " & errorMessage, _
                        tb.serviceProviders.nameFromHandle(spHandle)
Case TradeBuildSP.StandardSPErrorCodes.CIReConnectingDataSource
    tb.SetServiceProviderError errorCode, _
                        "Reconnecting to Contract database: " & errorMessage, _
                        tb.serviceProviders.nameFromHandle(spHandle)
Case TradeBuildSP.StandardSPErrorCodes.CIRetryConnectDataSource
    tb.SetServiceProviderError errorCode, _
                        "Attempting again to connect to Contract database: " & errorMessage, _
                        tb.serviceProviders.nameFromHandle(spHandle)
Case Else
    tb.SetServiceProviderError errorCode, _
                        errorMessage, _
                        nameFromHandle(spHandle)
End Select
End Sub

Private Sub IContractInfoServiceConsumer_LoadComplete( _
                            ByVal handle As Long, _
                            ByVal pContracts As TradeBuildSP.IContracts)
Dim lContracts As TradeBuild.Contracts

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

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get availableStudies() As StudyListEntry()
ReDim studylist(20) As StudyListEntry
Dim spList() As String
Dim nextStudyListIndex As Long
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long
Dim j As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.studies) <> 0 Then
            Set studySP = mServiceProviders(i).SPObject
            spList = studySP.implementedStudyNames
            For j = 0 To UBound(spList)
                If nextStudyListIndex > UBound(studylist) Then
                    ReDim studylist(UBound(studylist) + 20) As StudyListEntry
                End If
                studylist(nextStudyListIndex).name = spList(j)
                studylist(nextStudyListIndex).serviceProvider = mServiceProviders(i).name
                nextStudyListIndex = nextStudyListIndex + 1
            Next
        End If
    End If
Next

ReDim Preserve studylist(nextStudyListIndex - 1) As StudyListEntry
availableStudies = studylist
End Property

Friend Property Get studyDefinition( _
                ByVal name As String, _
                ByVal spName As String) As studyDefinition
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.studies) <> 0 And _
            (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set studyDefinition = studySP.studyDefinition(name)
            If Not studyDefinition Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Property

Friend Property Get SupportedInputTickfileFormats() As TradeBuild.TickfileFormatSpecifier()
SupportedInputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Replay)
End Property

Friend Property Get SupportedOutputTickfileFormats() As TradeBuild.TickfileFormatSpecifier()
SupportedOutputTickfileFormats = GetSupportedTickfileFormats(TradeBuildSP.TickfileServiceProviderCapabilities.Record)
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = New WeakReference
mTradeBuildAPIRef.Initialise value
End Property

'================================================================================
' Methods
'================================================================================

Public Function add( _
                            ByVal serviceProvider As TradeBuildSP.ICommonServiceProvider, _
                            Optional ByVal logLevel As TradeBuild.LogLevels, _
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
                    "TradeBuild.ServiceProviders::Add", _
                    "This Service Provider object has already been added"
        End If
        If mServiceProviders(i).name = serviceProvider.name Then
            err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "TradeBuild.ServiceProviders::Add", _
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
                            ByVal contractSpec As TradeBuild.contractSpecifier, _
                            ByVal historicDataPath As String _
                            ) As TradeBuild.HistoricDataWriter
                            
Dim histDataSp As TradeBuildSP.IBarDataServiceProvider
Dim histDataSPWriter As TradeBuildSP.IBarDataWriter
Dim histDataWriter As TradeBuild.HistoricDataWriter
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

Set histDataWriter = New TradeBuild.HistoricDataWriter
histDataWriter.ServiceProviderWriter = histDataSPWriter
Set CreateHistoricalDataWriter = histDataWriter

End Function

Friend Function CreateLiveOrderSubmitter( _
                            ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                            ByVal pContract As TradeBuild.Contract _
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
                                    ByVal pTicker As ticker, _
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
                            ByVal pContract As TradeBuild.Contract) As TradeBuildSP.IOrderSubmitter
                            
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
    Set lOrderSimulatorSp = New TradeBuild.OrderSimulatorSP
    add lOrderSimulatorSp
End If

Set orderSubmitter = lOrderSimulatorSp.CreateOrderSubmitter(serviceConsumer, pContract)

Set CreateSimulatedOrderSubmitter = orderSubmitter

                            
End Function

Friend Function CreateStudy( _
                ByVal studyName As String, _
                ByVal spName As String, _
                ByVal serviceConsumer As TradeBuildSP.IStudyServiceConsumer) As IStudy
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.studies) <> 0 And _
            (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set CreateStudy = studySP.CreateStudy(studyName, serviceConsumer)
            If Not CreateStudy Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

Friend Function CreateTickfileReader( _
                                    ByVal pTicker As ticker, _
                                    ByVal serviceConsumer As TradeBuildSP.ITickfileInputServiceConsumer, _
                                    ByVal inputTickfileFormatID As String _
                                    ) As TradeBuild.TickfileReader
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPReader As TradeBuildSP.ITickfileReader
Dim tfReader As TradeBuild.TickfileReader
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

Set tfReader = New TradeBuild.TickfileReader
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
                            ) As TradeBuild.TickfileWriter
                            
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim tickfileSPWriter As TradeBuildSP.ITickfileWriter
Dim tfWriter As TradeBuild.TickfileWriter
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

Set tfWriter = New TradeBuild.TickfileWriter
Set tickfileSPWriter = tickfileSp.CreateTickfileWriter(serviceConsumer, _
                                                    outputTickfileFormatID, _
                                                    outputTickfilePath)

tfWriter.ServiceProviderWriter = tickfileSPWriter

Set CreateTickfileWriter = tfWriter

End Function

Friend Function FetchStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal spName As String) As IParameters
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.studies) <> 0 And _
            (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            Set FetchStudyDefaultParameters = studySP.getStudyDefaultParameters(studyName)
            If Not FetchStudyDefaultParameters Is Nothing Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

Friend Sub LoadContracts( _
                            ByVal Contracts As TradeBuild.Contracts)

If Not LoadContractsFromNextSP(0, Contracts) Then Contracts.LoadComplete
End Sub

Friend Function nameFromHandle( _
                ByVal handle As Long) As String
If handle < 0 Then Exit Function
nameFromHandle = mServiceProviders(handle).name
End Function

Public Sub remove(ByVal sp As TradeBuildSP.ICommonServiceProvider)
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

Public Sub removeCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.remove i
Next
End Sub

Friend Function SetStudyDefaultParameters( _
                ByVal studyName As String, _
                ByVal spName As String, _
                ByVal params As IParameters) As Boolean
Dim studySP As TradeBuildSP.IStudyServiceProvider
Dim i As Long

For i = 0 To mMaxSPIndex
    If Not mServiceProviders(i).SPObject Is Nothing Then
        If (mServiceProviders(i).SPTypes And TradeBuildSP.ServiceProviderTypes.studies) <> 0 And _
            (mServiceProviders(i).name = spName Or _
            spName = "") _
        Then
            Set studySP = mServiceProviders(i).SPObject
            SetStudyDefaultParameters = studySP.SetStudyDefaultParameters(studyName, params)
            If SetStudyDefaultParameters Or _
                spName <> "" _
                Then Exit For
        End If
    End If
Next
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As TradeBuildSP.ICommonServiceProvider)
Dim listener As CollectionChangeListener
Dim i As Long
Dim Change As CollectionChangeEvent
Set Change.source = Me
Change.changeType = changeType
Set Change.affectedObject = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change Change
Next
End Sub

Private Function GetSupportedTickfileFormats( _
                            ByVal capabilities As TradeBuildSP.TickfileServiceProviderCapabilities _
                            ) As TradeBuild.TickfileFormatSpecifier()
Dim i As Long
Dim j As Long
Dim k As Long
Dim tickfileSp As TradeBuildSP.ITickfileServiceProvider
Dim supportedFormats() As TradeBuild.TickfileFormatSpecifier
ReDim allSupportedFormats(9) As TradeBuild.TickfileFormatSpecifier

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
                        ReDim Preserve allSupportedFormats(UBound(allSupportedFormats) + 10) As TradeBuild.TickfileFormatSpecifier
                    End If
                    allSupportedFormats(k) = supportedFormats(j)
                End If
            Next
        End If
    End If
Next

If k >= 0 Then
    ReDim Preserve allSupportedFormats(k) As TradeBuild.TickfileFormatSpecifier
    GetSupportedTickfileFormats = allSupportedFormats
Else
    Erase allSupportedFormats
End If
End Function


Private Function LoadContractsFromNextSP( _
                            ByVal startingIndex As Long, _
                            ByVal pContracts As TradeBuild.Contracts _
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
Set tb = mTradeBuildAPIRef.target
End Function
