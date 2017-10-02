VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarketDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type DataSourceBaseTableEntry
    InUse                   As Boolean
    Next                    As Long
    DataSourceBase          As DataSourceBase
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataManager"

Private Const ConfigSectionMarketDataSource         As String = "MarketDataSource"
Private Const ConfigSettingUseLocalTimeForStudies   As String = "&UseLocalTimeForStudies"

'@================================================================================
' Member variables
'@================================================================================

Private mFactory                                    As IMarketDataFactory
Attribute mFactory.VB_VarHelpID = -1

Private mStudyLibManager                            As StudyLibraryManager

Private mDataSourceBaseTable()                      As DataSourceBaseTableEntry
Private mNextDataSourceBaseTableIndex               As Long
Private mFirstFreeDataSourceBaseTableIndex          As Long
Private mLastFreeDataSourceBaseTableIndex           As Long

Private mContractMap                                As ContractMap

Private mConfig                                     As ConfigurationSection

Private mDataSources                                As New EnumerableCollection

Private mOptions                                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mDataSourceBaseTable(3) As DataSourceBaseTableEntry
mFirstFreeDataSourceBaseTableIndex = NullIndex
mLastFreeDataSourceBaseTableIndex = NullIndex
Set mContractMap = New ContractMap
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get DataSources() As EnumerableCollection
Set DataSources = mDataSources
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pKey As String, _
                ByVal pStateChangeListener As IStateChangeListener, _
                ByVal pIsRecoverable As Boolean, _
                ByVal pUseLocalTimeForStudies As Boolean, _
                ByVal pTickfileStreamId As Long) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture cannot be Nothing"

If pKey = "" Then pKey = GenerateGUIDString
Dim lDataSource As New MarketDataSource
If Not pStateChangeListener Is Nothing Then lDataSource.AddStateChangeListener pStateChangeListener

Dim lConfig As ConfigurationSection
If pIsRecoverable Then Set lConfig = getConfigSection(pKey, pUseLocalTimeForStudies)

Dim lStudyManager As StudyManager
If Not mStudyLibManager Is Nothing Then Set lStudyManager = mStudyLibManager.CreateStudyManagerWithSessionFuture(lDataSource.SessionFuture, pUseLocalTimeForStudies)

lDataSource.Initialise Me, pKey, pContractFuture, pClockFuture, lStudyManager, lConfig, pTickfileStreamId

mDataSources.Add lDataSource, pKey

Set CreateMarketDataSource = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim i As Long

For i = 0 To mNextDataSourceBaseTableIndex - 1
    If Not mDataSourceBaseTable(i).DataSourceBase Is Nothing Then
        mDataSourceBaseTable(i).DataSourceBase.Finish
        ReleaseDataSourceBase mDataSourceBaseTable(i).DataSourceBase
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetMarketDataSourceByKey(ByVal pKey As String) As IMarketDataSource
Const ProcName As String = "GetMarketDataSourceByKey"
On Error GoTo Err

If mDataSources.Contains(pKey) Then Set GetMarketDataSourceByKey = mDataSources.Item(pKey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetDataSourceBase( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                Optional ByVal pTickfileStreamId As Long = NullIndex) As DataSourceBase
Const ProcName As String = "GetDataSourceBase"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture cannot be Nothing"
AssertArgument pContractFuture.IsAvailable, "pContractFuture must be IsAvailable"
AssertArgument pTickfileStreamId = NullIndex Or pTickfileStreamId >= 0, "Invalid TickfileStreamId"

Dim lContract As IContract
Set lContract = pContractFuture.Value

Dim lIndex As Long
If pTickfileStreamId <> NullIndex Then
    lIndex = allocateDataSourceBaseTableEntry(pTickfileStreamId)
Else
    lIndex = getDataSourceBaseIndex(lContract.Specifier)
    If lIndex = -1 Then Exit Function
End If

Dim lDataSourceBase As DataSourceBase
Set lDataSourceBase = mDataSourceBaseTable(lIndex).DataSourceBase
If lDataSourceBase Is Nothing Then
    Set lDataSourceBase = New DataSourceBase
    Set mDataSourceBaseTable(lIndex).DataSourceBase = lDataSourceBase
    
    lDataSourceBase.Initialise Me, pContractFuture, pClockFuture, pTickfileStreamId <> NullIndex, lIndex, mOptions, mNumberOfMarketDepthRows
End If

Set GetDataSourceBase = lDataSourceBase

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pFactory As IMarketDataFactory, _
                ByVal pStudyLibManager As StudyLibraryManager, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mFactory = pFactory
Set mStudyLibManager = pStudyLibManager
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Assert Not pConfig Is Nothing, "pConfig Is Nothing"

gLogger.Log "Loading market data sources from config", ProcName, ModuleName

Set mConfig = pConfig

Dim lRemovableCSs As New Collection

Dim lCs As ConfigurationSection
For Each lCs In mConfig
    gLogger.Log "Loading market data source: " & lCs.InstanceQualifier, ProcName, ModuleName
    Dim lContractConfig As ConfigurationSection
    Set lContractConfig = lCs.GetConfigurationSection(ConfigSectionContract)
    If lContractConfig Is Nothing Then
        lRemovableCSs.Add lCs
    Else
        Dim lContractFuture As IFuture
        Set lContractFuture = CreateFuture(LoadContractFromConfig(lContractConfig))
        CreateMarketDataSource lContractFuture, _
                                CreateClockFuture(lContractFuture), _
                                lCs.InstanceQualifier, _
                                Nothing, _
                                True, _
                                lCs.GetSetting(ConfigSettingUseLocalTimeForStudies, "False"), _
                                NullIndex
    End If
Next

For Each lCs In lRemovableCSs
    lCs.Remove
Next

gLogger.Log "Loaded market data sources", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ReleaseDataSourceBase(ByVal pDataSourceBase As DataSourceBase)
Const ProcName As String = "ReleaseDataSourceBase"
On Error GoTo Err

releaseDataSourceBaseTableEntry pDataSourceBase.StreamId
Dim lContract As IContract
Set lContract = pDataSourceBase.ContractFuture.Value
If mContractMap.Contains(lContract.Specifier) Then mContractMap.Remove lContract.Specifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveMarketDataSource(ByVal pDataSource As MarketDataSource)
Const ProcName As String = "RemoveMarketDataSource"
On Error GoTo Err

mDataSources.Remove pDataSource.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RequestMarketData(ByVal pDataSourceBase As DataSourceBase) As ITickStream
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Set RequestMarketData = mFactory.CreateMarketDataStream(pDataSourceBase.ContractFuture, pDataSourceBase, pDataSourceBase.StreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function RequestMarketDepth(ByVal pDataSourceBase As DataSourceBase, ByVal pNumberOfRows As Long) As ITickStream
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Set RequestMarketDepth = mFactory.CreateMarketDepthStream(pDataSourceBase.ContractFuture, pNumberOfRows, pDataSourceBase, pDataSourceBase.StreamId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateDataSourceBaseTableEntry(Optional ByVal pIndex As Long = NullIndex) As Long
Const ProcName As String = "allocateDataSourceBaseTableEntry"
On Error GoTo Err

If pIndex <> NullIndex Then
    ensureTableSpace pIndex
    allocateDataSourceBaseTableEntry = pIndex
ElseIf findNextFreeDataSourceBasetableEntry Then
    allocateDataSourceBaseTableEntry = mNextDataSourceBaseTableIndex
ElseIf mFirstFreeDataSourceBaseTableIndex <> NullIndex Then
    allocateDataSourceBaseTableEntry = mFirstFreeDataSourceBaseTableIndex
    mFirstFreeDataSourceBaseTableIndex = mDataSourceBaseTable(mFirstFreeDataSourceBaseTableIndex).Next
    If mFirstFreeDataSourceBaseTableIndex = NullIndex Then mLastFreeDataSourceBaseTableIndex = NullIndex
Else
    ensureTableSpace mNextDataSourceBaseTableIndex
    allocateDataSourceBaseTableEntry = mNextDataSourceBaseTableIndex
End If

Assert Not (pIndex = NullIndex And mDataSourceBaseTable(allocateDataSourceBaseTableEntry).InUse), "Stream table entry already in use"
Assert allocateDataSourceBaseTableEntry <= mFactory.MaxStreamId, "Max stream id exceeded"

mDataSourceBaseTable(allocateDataSourceBaseTableEntry).InUse = True
mDataSourceBaseTable(allocateDataSourceBaseTableEntry).Next = NullIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ensureTableSpace(ByVal pIndex As Long)
Const ProcName As String = "ensureTableSpace"
On Error GoTo Err

Do While pIndex > UBound(mDataSourceBaseTable)
    ReDim Preserve mDataSourceBaseTable(2 * (UBound(mDataSourceBaseTable) + 1) - 1) As DataSourceBaseTableEntry
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findNextFreeDataSourceBasetableEntry() As Boolean
Const ProcName As String = "findNextFreeDataSourceBasetableEntry"
On Error GoTo Err

Dim lFound As Boolean
lFound = False

Do While mNextDataSourceBaseTableIndex <= UBound(mDataSourceBaseTable)
    If Not mDataSourceBaseTable(mNextDataSourceBaseTableIndex).InUse Then
        lFound = True
        Exit Do
    End If
    mNextDataSourceBaseTableIndex = mNextDataSourceBaseTableIndex + 1
Loop

If mNextDataSourceBaseTableIndex > mFactory.MaxStreamId Then lFound = False
findNextFreeDataSourceBasetableEntry = lFound

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigSection( _
                ByVal pKey As String, _
                ByVal pUseLocalTimeForStudies As Boolean) As ConfigurationSection
Const ProcName As String = "getConfigSection"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function

Dim lConfig As ConfigurationSection
Set lConfig = mConfig.AddConfigurationSection(ConfigSectionMarketDataSource & "(" & pKey & ")")
lConfig.SetSetting ConfigSettingUseLocalTimeForStudies, pUseLocalTimeForStudies

Set getConfigSection = lConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDataSourceBaseIndex( _
                ByVal pContractSpec As IContractSpecifier) As Long
Const ProcName As String = "getDataSourceBaseIndex"
On Error GoTo Err

Dim lIndex As Long

If mContractMap.Contains(pContractSpec) Then
    lIndex = mContractMap.Item(pContractSpec)
Else
    lIndex = allocateDataSourceBaseTableEntry
    mContractMap.Add pContractSpec, lIndex
End If

getDataSourceBaseIndex = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseDataSourceBaseTableEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseDataSourceBaseTableEntry"
On Error GoTo Err

If mLastFreeDataSourceBaseTableIndex <> NullIndex Then
    mDataSourceBaseTable(mLastFreeDataSourceBaseTableIndex).Next = pIndex
Else
    mFirstFreeDataSourceBaseTableIndex = pIndex
End If
mLastFreeDataSourceBaseTableIndex = pIndex

mDataSourceBaseTable(pIndex).Next = NullIndex
mDataSourceBaseTable(pIndex).InUse = False
Set mDataSourceBaseTable(pIndex).DataSourceBase = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



