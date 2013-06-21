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

Implements TaskCompletionListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type StreamTableEntry
    Next                    As Long
    MarketDataStream        As MarketDataStream
    TickStream              As ITickStream
    DepthStream             As ITickStream
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MarketDataManager"

Private Const ConfigSectionMarketDataSource         As String = "MarketDataSource"

Private Const ConfigSettingKey                      As String = "&Key"

'@================================================================================
' Member variables
'@================================================================================

Private mFactory                                    As IMarketDataFactory
Attribute mFactory.VB_VarHelpID = -1
'Private mPrimaryContractStore                       As IContractStore
'Private mSecondaryContractStore                     As IContractStore

Private mStreamTable()                              As StreamTableEntry
Private mNextStreamTableIndex                       As Long
Private mFirstFreeStreamTableIndex                  As Long
Private mLastFreeStreamTableIndex                   As Long

Private mStreamMap                                  As ContractStreamMap

Private mConfig                                     As ConfigurationSection

Private mDataSources                                As New EnumerableCollection

Private mOptions                                    As MarketDataSourceOptions
Private mNumberOfMarketDepthRows                    As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStreamTable(3) As StreamTableEntry
mFirstFreeStreamTableIndex = NullIndex
mLastFreeStreamTableIndex = NullIndex
Set mStreamMap = New ContractStreamMap
End Sub

'@================================================================================
' TaskCompletionListener Interface Members
'@================================================================================

Private Sub TaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_TaskCompleted"
On Error GoTo Err

Dim lIndex As Long
lIndex = ev.Cookie

If ev.Cancelled Then
    mStreamTable(lIndex).MarketDataStream.NotifyError MarketDataErrorCodes.MarketDataErrCancelled, "The contract fetch operation was cancelled"
ElseIf ev.ErrorNumber <> 0 Then
    mStreamTable(lIndex).MarketDataStream.NotifyError MarketDataErrorCodes.MarketDataErrContractFetchError, ev.ErrorMessage
Else
    Dim lContracts As IContracts
    Set lContracts = ev.Result
    
    If lContracts.Count = 0 Then
        mStreamTable(lIndex).MarketDataStream.NotifyError MarketDataErrorCodes.MarketDataErrNoContract, "No such contract exists"
    ElseIf lContracts.Count > 1 Then
        mStreamTable(lIndex).MarketDataStream.NotifyError MarketDataErrorCodes.MarketDataErrContractNotUnique, "Contract not uniquely specified"
    Else
        SetContract lIndex, lContracts(1)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get DataSources() As EnumerableCollection
Set DataSources = mDataSources
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub cancelMarketData(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "CancelMarketData"
On Error GoTo Err

Dim lIndex As Long

lIndex = pDataStream.StreamId

If Not mStreamTable(lIndex).TickStream Is Nothing Then mStreamTable(lIndex).TickStream.Finish

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub cancelMarketDepth(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "CancelMarketDepth"
On Error GoTo Err

Dim lIndex As Long

lIndex = pDataStream.StreamId

If Not mStreamTable(lIndex).DepthStream Is Nothing Then mStreamTable(lIndex).DepthStream.Finish

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CreateMarketDataSource( _
                ByVal pContract As IContract, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pKey As String, _
                ByVal pStateChangeListener As StateChangeListener, _
                Optional ByVal pIndexToUse As Long = -1, _
                Optional ByVal pUseExistingDataStream As Boolean) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

AssertArgument Not (pContract Is Nothing And pContractSpec Is Nothing), "Either pContract or pContract spec must be supplied"
AssertArgument pContract Is Nothing Xor pContractSpec Is Nothing, "Either pContract or pContract spec must be supplied"

Dim lContractSpec As IContractSpecifier
If Not pContract Is Nothing Then
    Set lContractSpec = pContract.Specifier
Else
    Set lContractSpec = pContractSpec
End If

gLogger.Log "Creating market data source for", ProcName, ModuleName, , lContractSpec.ToString

Dim lIndex As Long
lIndex = getDataStreamIndex(lContractSpec, pIndexToUse, pUseExistingDataStream)
If lIndex = -1 Then Exit Function

Dim lDataStream As MarketDataStream
Set lDataStream = mStreamTable(lIndex).MarketDataStream
If lDataStream Is Nothing Then
    Set lDataStream = New MarketDataStream
    Set mStreamTable(lIndex).MarketDataStream = lDataStream
    lDataStream.Initialise Me, pContract, lContractSpec, lIndex, mOptions, mNumberOfMarketDepthRows
ElseIf Not lDataStream.Contract Is Nothing Then
ElseIf Not pContract Is Nothing Then
    lDataStream.Contract = pContract
End If

If pKey = "" Then pKey = GenerateGUIDString
Dim lDataSource As New MarketDataSource
If Not pStateChangeListener Is Nothing Then lDataSource.AddStateChangeListener pStateChangeListener
lDataSource.Initialise pKey, lDataStream, getConfigSection(pContract, mOptions, pKey, mNumberOfMarketDepthRows)

lDataStream.AddDataSource lDataSource

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

For i = 0 To mNextStreamTableIndex - 1
    If Not mStreamTable(i).MarketDataStream Is Nothing Then
        If Not mStreamTable(i).TickStream Is Nothing Then mStreamTable(i).TickStream.Finish
        If Not mStreamTable(i).DepthStream Is Nothing Then mStreamTable(i).DepthStream.Finish
        ReleaseStream mStreamTable(i).MarketDataStream
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

Friend Sub Initialise( _
                ByVal pFactory As IMarketDataFactory, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long, _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mFactory = pFactory
mOptions = pOptions
mNumberOfMarketDepthRows = pNumberOfMarketDepthRows

Set mConfig = pConfig
If Not mConfig Is Nothing Then LoadFromConfig mConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Dim lCs As ConfigurationSection
Dim lDataSource As IMarketDataSource

For Each lCs In mConfig
    Dim lContractConfig As ConfigurationSection
    Set lContractConfig = lCs.GetConfigurationSection(ConfigSectionContract)
    If Not lContractConfig Is Nothing Then
        CreateMarketDataSource LoadContractFromConfig(lContractConfig), _
                                Nothing, _
                                lCs.GetSetting(ConfigSettingKey), _
                                Nothing
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyError"
On Error GoTo Err

mStreamTable(pStreamId).MarketDataStream.NotifyError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDataError"
On Error GoTo Err

mStreamTable(pStreamId).MarketDataStream.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "NotifyMarketDepthError"
On Error GoTo Err

mStreamTable(pStreamId).MarketDataStream.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "NotifyTick"
On Error GoTo Err

mStreamTable(ev.StreamId).MarketDataStream.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ReleaseStream(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "ReleaseStream"
On Error GoTo Err

releaseStreamTableEntry pDataStream.StreamId
If Not pDataStream.Contract Is Nothing Then mStreamMap.Remove pDataStream.Contract.Specifier

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
If Not mConfig Is Nothing Then mConfig.RemoveConfigurationSection ConfigSectionMarketDataSource & "(" & pDataSource.Key & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub requestMarketData(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Assert mStreamTable(pDataStream.StreamId).TickStream Is Nothing, "Market data has already been requested"
Set mStreamTable(pDataStream.StreamId).TickStream = mFactory.CreateMarketDataStream(pDataStream.Contract, Me, Me, pDataStream.StreamId)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub requestMarketDepth(ByVal pDataStream As MarketDataStream, ByVal pNumberOfRows As Long)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Assert mStreamTable(pDataStream.StreamId).DepthStream Is Nothing, "Market depth has already been requested"
Set mStreamTable(pDataStream.StreamId).DepthStream = mFactory.CreateMarketDepthStream(pDataStream.Contract, pNumberOfRows, Me, Me, pDataStream.StreamId)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetContract(ByVal pStreamId As Long, pContract As IContract)
Const ProcName As String = "SetContract"
On Error GoTo Err

If mStreamTable(pStreamId).MarketDataStream.Contract Is Nothing Then mStreamTable(pStreamId).MarketDataStream.Contract = pContract

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateStreamTableEntry() As Long
Const ProcName As String = "allocateStreamTableEntry"
On Error GoTo Err

If mNextStreamTableIndex <= UBound(mStreamTable) Then
    allocateStreamTableEntry = mNextStreamTableIndex
    mNextStreamTableIndex = mNextStreamTableIndex + 1
ElseIf mFirstFreeStreamTableIndex <> NullIndex Then
    allocateStreamTableEntry = mFirstFreeStreamTableIndex
    mFirstFreeStreamTableIndex = mStreamTable(mFirstFreeStreamTableIndex).Next
    If mFirstFreeStreamTableIndex = NullIndex Then mLastFreeStreamTableIndex = NullIndex
Else
    Dim lNewBound As Long
    Assert UBound(mStreamTable) < mFactory.MaxStreamId, "Cannot allocate Stream table entry - table is full"
    lNewBound = 2 * (UBound(mStreamTable) + 1) - 1
    If lNewBound > mFactory.MaxStreamId Then lNewBound = mFactory.MaxStreamId
    ReDim Preserve mStreamTable(lNewBound) As StreamTableEntry
    allocateStreamTableEntry = mNextStreamTableIndex
    mNextStreamTableIndex = mNextStreamTableIndex + 1
End If

mStreamTable(allocateStreamTableEntry).Next = NullIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigSection( _
                Optional ByVal pContract As IContract, _
                Optional ByVal pOptions As MarketDataSourceOptions = MarketDataSourceOptUseExchangeTimeZone, _
                Optional ByVal pKey As String, _
                Optional ByVal pNumberOfMarketDepthRows As Long) As ConfigurationSection
Const ProcName As String = "getConfigSection"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function


Dim lConfig As ConfigurationSection
Set lConfig = mConfig.AddConfigurationSection(ConfigSectionMarketDataSource & "(" & pKey & ")")

If Not pContract Is Nothing Then SaveContractToConfig pContract, lConfig.AddConfigurationSection(ConfigSectionContract)
lConfig.SetSetting ConfigSettingKey, pKey

Set getConfigSection = lConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDataStreamIndex( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIndexToUse As Long, _
                ByVal pUseExistingDataStream As Boolean) As Long
Const ProcName As String = "getDataStreamIndex"
On Error GoTo Err

Dim lIndex As Long

If mStreamMap.Contains(pContractSpec) Then
    lIndex = mStreamMap.Item(pContractSpec)
ElseIf pIndexToUse >= 0 Then
    lIndex = pIndexToUse
    mStreamMap.Add pContractSpec, lIndex
ElseIf pUseExistingDataStream Then
    lIndex = -1
Else
    lIndex = allocateStreamTableEntry
    mStreamMap.Add pContractSpec, lIndex
End If

getDataStreamIndex = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseStreamTableEntry(ByVal pindex As Long)
Const ProcName As String = "releaseStreamTableEntry"

On Error GoTo Err

If mLastFreeStreamTableIndex <> NullIndex Then
    mStreamTable(mLastFreeStreamTableIndex).Next = pindex
Else
    mFirstFreeStreamTableIndex = pindex
End If
mLastFreeStreamTableIndex = pindex

mStreamTable(pindex).Next = NullIndex
Set mStreamTable(pindex).MarketDataStream = Nothing
Set mStreamTable(pindex).TickStream = Nothing
Set mStreamTable(pindex).DepthStream = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



