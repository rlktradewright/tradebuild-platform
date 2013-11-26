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

Implements IGenericTickListener
Implements IMarketDataErrorListener
Implements StateChangeListener

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
    InUse                   As Boolean
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

'@================================================================================
' Member variables
'@================================================================================

Private mFactory                                    As IMarketDataFactory
Attribute mFactory.VB_VarHelpID = -1

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
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

mStreamTable(ev.StreamId).MarketDataStream.NotifyTick ev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IMarketDataErrorListener Interface Members
'@================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDataError"
On Error GoTo Err

mStreamTable(pStreamId).MarketDataStream.NotifyMarketDataError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDepthError"
On Error GoTo Err

mStreamTable(pStreamId).MarketDataStream.NotifyMarketDepthError pErrorCode, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

If ev.State <> TickStreamStates.TickStreamStateFinished Then Exit Sub

Dim lStream As TickStream
Set lStream = ev.Source
mStreamTable(lStream.StreamId).MarketDataStream.NotifyStreamFinished

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get DataSources() As EnumerableCollection
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

If Not mStreamTable(lIndex).TickStream Is Nothing Then
    mStreamTable(lIndex).TickStream.Finish
    Set mStreamTable(lIndex).TickStream = Nothing
End If

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

If Not mStreamTable(lIndex).DepthStream Is Nothing Then
    mStreamTable(lIndex).DepthStream.Finish
    Set mStreamTable(lIndex).DepthStream = Nothing
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pKey As String, _
                ByVal pStateChangeListener As StateChangeListener, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pIndexToUse As Long = NullIndex, _
                Optional ByVal pUseExistingDataStream As Boolean) As IMarketDataSource
Const ProcName As String = "CreateMarketDataSource"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture cannot be Nothing"

If pKey = "" Then pKey = GenerateGUIDString
Dim lDataSource As New MarketDataSource
If Not pStateChangeListener Is Nothing Then lDataSource.AddStateChangeListener pStateChangeListener

Dim lConfig As ConfigurationSection
If pIsRecoverable Then Set lConfig = getConfigSection(pKey)
lDataSource.Initialise Me, pKey, pContractFuture, mOptions, pIndexToUse, pUseExistingDataStream, lConfig

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

Friend Function GetMarketDataStream( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pIndexToUse As Long = -1, _
                Optional ByVal pUseExistingDataStream As Boolean, _
                Optional ByVal pSimulatedClock As Clock) As MarketDataStream
Const ProcName As String = "GetMarketDataStream"
On Error GoTo Err

AssertArgument Not pContractFuture Is Nothing, "pContractFuture cannot be Nothing"
AssertArgument pContractFuture.IsAvailable, "pContractFuture must be IsAvailable"

Dim lIndex As Long
Dim lContract As IContract
Set lContract = pContractFuture.Value
lIndex = getDataStreamIndex(lContract.Specifier, pIndexToUse, pUseExistingDataStream)
If lIndex = -1 Then Exit Function

Dim lDataStream As MarketDataStream
Set lDataStream = mStreamTable(lIndex).MarketDataStream
If lDataStream Is Nothing Then
    Set lDataStream = New MarketDataStream
    Set mStreamTable(lIndex).MarketDataStream = lDataStream
    
    Dim lClock As Clock
    If Not pSimulatedClock Is Nothing Then
        Set lClock = pSimulatedClock
    ElseIf mOptions And MarketDataSourceOptUseExchangeTimeZone Then
        Set lClock = GetClock(lContract.TimezoneName)
    Else
        Set lClock = GetClock("")
    End If

    lDataStream.Initialise Me, pContractFuture, lIndex, mOptions, mNumberOfMarketDepthRows, lClock
End If

Set GetMarketDataStream = lDataStream

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pFactory As IMarketDataFactory, _
                ByVal pOptions As MarketDataSourceOptions, _
                ByVal pNumberOfMarketDepthRows As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mFactory = pFactory
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

Dim lCs As ConfigurationSection
For Each lCs In mConfig
    gLogger.Log "Loading market data source: " & lCs.InstanceQualifier, ProcName, ModuleName
    Dim lContractConfig As ConfigurationSection
    Set lContractConfig = lCs.GetConfigurationSection(ConfigSectionContract)
    If Not lContractConfig Is Nothing Then
        CreateMarketDataSource CreateFuture(LoadContractFromConfig(lContractConfig)), _
                                lCs.InstanceQualifier, _
                                Nothing, _
                                True
    End If
Next

gLogger.Log "Loaded market data sources", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ReleaseStream(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "ReleaseStream"
On Error GoTo Err

releaseStreamTableEntry pDataStream.StreamId
Dim lContract As IContract
Set lContract = pDataStream.ContractFuture.Value
mStreamMap.Remove lContract.Specifier

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

Friend Sub RequestMarketData(ByVal pDataStream As MarketDataStream)
Const ProcName As String = "RequestMarketData"
On Error GoTo Err

Assert mStreamTable(pDataStream.StreamId).TickStream Is Nothing, "Market data has already been requested"
Set mStreamTable(pDataStream.StreamId).TickStream = mFactory.CreateMarketDataStream(pDataStream.ContractFuture, Me, pDataStream.StreamId)
mStreamTable(pDataStream.StreamId).TickStream.AddGenericTickListener Me
mStreamTable(pDataStream.StreamId).TickStream.AddStateChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub requestMarketDepth(ByVal pDataStream As MarketDataStream, ByVal pNumberOfRows As Long)
Const ProcName As String = "RequestMarketDepth"
On Error GoTo Err

Assert mStreamTable(pDataStream.StreamId).DepthStream Is Nothing, "Market depth has already been requested"
Set mStreamTable(pDataStream.StreamId).DepthStream = mFactory.CreateMarketDepthStream(pDataStream.ContractFuture, pNumberOfRows, Me, pDataStream.StreamId)
mStreamTable(pDataStream.StreamId).DepthStream.AddGenericTickListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateStreamTableEntry(Optional ByVal pIndex As Long = NullIndex) As Long
Const ProcName As String = "allocateStreamTableEntry"
On Error GoTo Err

If pIndex <> NullIndex Then
    ensureTableSpace pIndex
    allocateStreamTableEntry = pIndex
ElseIf findNextFreeEntryInCurrentSpace Then
    allocateStreamTableEntry = mNextStreamTableIndex
ElseIf mFirstFreeStreamTableIndex <> NullIndex Then
    allocateStreamTableEntry = mFirstFreeStreamTableIndex
    mFirstFreeStreamTableIndex = mStreamTable(mFirstFreeStreamTableIndex).Next
    If mFirstFreeStreamTableIndex = NullIndex Then mLastFreeStreamTableIndex = NullIndex
Else
    ensureTableSpace mNextStreamTableIndex
    allocateStreamTableEntry = mNextStreamTableIndex
End If

Assert Not mStreamTable(allocateStreamTableEntry).InUse, "Stream table entry already in use"
Assert allocateStreamTableEntry <= mFactory.MaxStreamId, "Max stream id exceeded"

mStreamTable(allocateStreamTableEntry).InUse = True
mStreamTable(allocateStreamTableEntry).Next = NullIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ensureTableSpace(ByVal pIndex As Long)
Const ProcName As String = "ensureTableSpace"
On Error GoTo Err

Do While pIndex > UBound(mStreamTable)
    ReDim Preserve mStreamTable(2 * (UBound(mStreamTable) + 1) - 1) As StreamTableEntry
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findNextFreeEntryInCurrentSpace() As Boolean
Const ProcName As String = "findNextFreeEntryInCurrentSpace"
On Error GoTo Err

Dim lFound As Boolean
lFound = False

Do While mNextStreamTableIndex <= UBound(mStreamTable)
    If Not mStreamTable(mNextStreamTableIndex).InUse Then
        lFound = True
        Exit Do
    End If
    mNextStreamTableIndex = mNextStreamTableIndex + 1
Loop

If mNextStreamTableIndex > mFactory.MaxStreamId Then lFound = False
findNextFreeEntryInCurrentSpace = lFound

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getConfigSection( _
                ByVal pKey As String) As ConfigurationSection
Const ProcName As String = "getConfigSection"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function

Dim lConfig As ConfigurationSection
Set lConfig = mConfig.AddConfigurationSection(ConfigSectionMarketDataSource & "(" & pKey & ")")

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
    lIndex = allocateStreamTableEntry(pIndexToUse)
    mStreamMap.Add pContractSpec, lIndex
ElseIf pUseExistingDataStream Then
    lIndex = NullIndex
Else
    lIndex = allocateStreamTableEntry
    mStreamMap.Add pContractSpec, lIndex
End If

getDataStreamIndex = lIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseStreamTableEntry(ByVal pIndex As Long)
Const ProcName As String = "releaseStreamTableEntry"

On Error GoTo Err

If mLastFreeStreamTableIndex <> NullIndex Then
    mStreamTable(mLastFreeStreamTableIndex).Next = pIndex
Else
    mFirstFreeStreamTableIndex = pIndex
End If
mLastFreeStreamTableIndex = pIndex

mStreamTable(pIndex).Next = NullIndex
mStreamTable(pIndex).InUse = False
Set mStreamTable(pIndex).MarketDataStream = Nothing
Set mStreamTable(pIndex).TickStream = Nothing
Set mStreamTable(pIndex).DepthStream = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



