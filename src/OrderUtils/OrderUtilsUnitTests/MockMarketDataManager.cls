VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockMarketDataManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements IMarketDataManager
Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type ReplayableTick
    Tick                As GenericTick
    DataSource          As MockDataSource
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MockMarketDataManager"

'@================================================================================
' Member variables
'@================================================================================

Private mRefreshInterval                            As Long

Private mDataSources                                As New Collection

Private mTicks                                      As SortedDictionary

Private mTaskContext                                As TaskContext

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTicks = CreateSortedDictionary(KeyTypeDate)
StartTask Me, PriorityNormal
End Sub

'@================================================================================
' IMarketDataManager Interface Members
'@================================================================================

Private Property Get IMarketDataManager_Capabilities() As MarketDataManagerCapabilities
IMarketDataManager_Capabilities = MarketDataManagerSnapshots Or MarketDataManagerStreaming
End Property

Private Function IMarketDataManager_CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener) As IMarketDataSource
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Function

Private Function IMarketDataManager_CreateSnapshotMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pKey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As IStateChangeListener, _
                Optional ByVal pRefreshInterval As Long = 0&) As IMarketDataSource
Dim lDataSource As New MockDataSource
lDataSource.Initialise pContractFuture, Me, getNextStreamId
mRefreshInterval = pRefreshInterval
mDataSources.Add lDataSource

Set IMarketDataManager_CreateSnapshotMarketDataSource = lDataSource
End Function

Private Property Get IMarketDataManager_DataSources() As EnumerableCollection
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Sub IMarketDataManager_Finish()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Function IMarketDataManager_GetMarketDataSource( _
                ByVal pKey As String) As IMarketDataSource
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Function

Private Sub IMarketDataManager_LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataManager_RemoveFromConfig()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Function IMarketDataManager_Supports( _
                ByVal pCapability As MarketDataManagerCapabilities) As Boolean
IMarketDataManager_Supports = True
End Function

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()
End Sub

Private Sub ITask_Run()
Dim en As Enumerator
Set en = mTicks.Enumerator

If en.MoveNext Then
    Dim lTick As ReplayableTick: lTick = en.Current
    Dim lDataSource As MockDataSource: Set lDataSource = lTick.DataSource
    lDataSource.FireTick lTick.Tick
    en.Remove
    mTaskContext.Suspend 5
Else
    mTaskContext.Suspend -1
End If
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function PlayTicks( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pDataSource As MockDataSource)
Const OneMilliSec As Double = 1# / 86400# / 1000#

Dim lTimestamp1 As Date: lTimestamp1 = CDate("2020-07-29 14:43:15")
Dim lTimestamp2 As Date: lTimestamp2 = CDate("2020-07-29 14:43:16")

Select Case pContractSpec.localSymbol
Case "MSFT"
    addTick pDataSource, lTimestamp1, TickTypeAsk, 202.65
    addTick pDataSource, lTimestamp1 + 2# * OneMilliSec, TickTypeBid, 202.62
    addTick pDataSource, lTimestamp1 + 3# * OneMilliSec, TickTypeTrade, 202.64
    addTick pDataSource, lTimestamp1 + 4# * OneMilliSec, -1, 0#
Case "MSFT  200731C00187500"
    addTick pDataSource, lTimestamp2, TickTypeAsk, 14.2
    addTick pDataSource, lTimestamp2 + 15# * OneMilliSec, TickTypeBid, 14.05
    addTick pDataSource, lTimestamp2 + 30# * OneMilliSec, TickTypeTrade, 14.15
    addTick pDataSource, lTimestamp2 + 45# * OneMilliSec, TickTypeOptionModelPrice, 14.08
    addTick pDataSource, lTimestamp2 + 60# * OneMilliSec, TickTypeOptionModelDelta, 0.99
    addTick pDataSource, lTimestamp2 + 75# * OneMilliSec, -1, 0#
Case "MSFT  200731C00190000"
    addTick pDataSource, lTimestamp2 + OneMilliSec, TickTypeAsk, 12.2
    addTick pDataSource, lTimestamp2 + 16# * OneMilliSec, TickTypeBid, 12.05
    addTick pDataSource, lTimestamp2 + 31# * OneMilliSec, TickTypeTrade, 12.15
    addTick pDataSource, lTimestamp2 + 46# * OneMilliSec, TickTypeOptionModelPrice, 12.08
    addTick pDataSource, lTimestamp2 + 61# * OneMilliSec, TickTypeOptionModelDelta, 0.95
    addTick pDataSource, lTimestamp2 + 76# * OneMilliSec, -1, 0#
Case "MSFT  200731C00192500"
    addTick pDataSource, lTimestamp2 + 2# * OneMilliSec, TickTypeAsk, 10.2
    addTick pDataSource, lTimestamp2 + 17# * OneMilliSec, TickTypeBid, 10.05
    addTick pDataSource, lTimestamp2 + 32# * OneMilliSec, TickTypeTrade, 10.15
    addTick pDataSource, lTimestamp2 + 47# * OneMilliSec, TickTypeOptionModelPrice, 10.1
    addTick pDataSource, lTimestamp2 + 62# * OneMilliSec, TickTypeOptionModelDelta, 0.85
    addTick pDataSource, lTimestamp2 + 77# * OneMilliSec, -1, 0#
Case "MSFT  200731C00195000"
    addTick pDataSource, lTimestamp2 + 3# * OneMilliSec, TickTypeAsk, 7.85
    addTick pDataSource, lTimestamp2 + 18# * OneMilliSec, TickTypeBid, 7.7
    addTick pDataSource, lTimestamp2 + 33# * OneMilliSec, TickTypeTrade, 7.8
    addTick pDataSource, lTimestamp2 + 48# * OneMilliSec, TickTypeOptionModelPrice, 7.8
    addTick pDataSource, lTimestamp2 + 63# * OneMilliSec, TickTypeOptionModelDelta, 0.8
    addTick pDataSource, lTimestamp2 + 78# * OneMilliSec, -1, 0#
Case "MSFT  200731C00197500"
    addTick pDataSource, lTimestamp2 + 4# * OneMilliSec, TickTypeAsk, 5.55
    addTick pDataSource, lTimestamp2 + 19# * OneMilliSec, TickTypeBid, 5.48
    addTick pDataSource, lTimestamp2 + 34# * OneMilliSec, TickTypeTrade, 5.49
    addTick pDataSource, lTimestamp2 + 49# * OneMilliSec, TickTypeOptionModelPrice, 5.51
    addTick pDataSource, lTimestamp2 + 64# * OneMilliSec, TickTypeOptionModelDelta, 0.65
    addTick pDataSource, lTimestamp2 + 79# * OneMilliSec, -1, 0#
Case "MSFT  200731C00200000"
    addTick pDataSource, lTimestamp2 + 5# * OneMilliSec, TickTypeAsk, 3.48
    addTick pDataSource, lTimestamp2 + 20# * OneMilliSec, TickTypeBid, 3.44
    addTick pDataSource, lTimestamp2 + 35# * OneMilliSec, TickTypeTrade, 3.44
    addTick pDataSource, lTimestamp2 + 50# * OneMilliSec, TickTypeOptionModelPrice, 3.46
    addTick pDataSource, lTimestamp2 + 65# * OneMilliSec, TickTypeOptionModelDelta, 0.65
    addTick pDataSource, lTimestamp2 + 80# * OneMilliSec, -1, 0#
Case "MSFT  200731C00202500"
    addTick pDataSource, lTimestamp2 + 6# * OneMilliSec, TickTypeAsk, 1.88
    addTick pDataSource, lTimestamp2 + 21# * OneMilliSec, TickTypeBid, 1.85
    addTick pDataSource, lTimestamp2 + 36# * OneMilliSec, TickTypeTrade, 1.88
    addTick pDataSource, lTimestamp2 + 51# * OneMilliSec, TickTypeOptionModelPrice, 1.87
    addTick pDataSource, lTimestamp2 + 66# * OneMilliSec, TickTypeOptionModelDelta, 0.55
    addTick pDataSource, lTimestamp2 + 81# * OneMilliSec, -1, 0#
Case "MSFT  200731C00205000"
    addTick pDataSource, lTimestamp2 + 7# * OneMilliSec, TickTypeAsk, 0.9
    addTick pDataSource, lTimestamp2 + 22# * OneMilliSec, TickTypeBid, 0.86
    addTick pDataSource, lTimestamp2 + 37# * OneMilliSec, TickTypeTrade, 0.88
    addTick pDataSource, lTimestamp2 + 52# * OneMilliSec, TickTypeOptionModelPrice, 0.89
    addTick pDataSource, lTimestamp2 + 67# * OneMilliSec, TickTypeOptionModelDelta, 0.47
    addTick pDataSource, lTimestamp2 + 82# * OneMilliSec, -1, 0#
Case "MSFT  200731C00207500"
    addTick pDataSource, lTimestamp2 + 8# * OneMilliSec, TickTypeAsk, 0.4
    addTick pDataSource, lTimestamp2 + 23# * OneMilliSec, TickTypeBid, 0.36
    addTick pDataSource, lTimestamp2 + 38# * OneMilliSec, TickTypeTrade, 0.38
    addTick pDataSource, lTimestamp2 + 53# * OneMilliSec, TickTypeOptionModelPrice, 0.39
    addTick pDataSource, lTimestamp2 + 68# * OneMilliSec, TickTypeOptionModelDelta, 0.35
    addTick pDataSource, lTimestamp2 + 83# * OneMilliSec, -1, 0#
Case "MSFT  200731C00210000"
    addTick pDataSource, lTimestamp2 + 9# * OneMilliSec, TickTypeAsk, 0.2
    addTick pDataSource, lTimestamp2 + 24# * OneMilliSec, TickTypeBid, 0.16
    addTick pDataSource, lTimestamp2 + 39# * OneMilliSec, TickTypeTrade, 0.18
    addTick pDataSource, lTimestamp2 + 54# * OneMilliSec, TickTypeOptionModelPrice, 0.17
    addTick pDataSource, lTimestamp2 + 69# * OneMilliSec, TickTypeOptionModelDelta, 0.25
    addTick pDataSource, lTimestamp2 + 84# * OneMilliSec, -1, 0#
Case "MSFT  200731C00212500"
    addTick pDataSource, lTimestamp2 + 10# * OneMilliSec, TickTypeAsk, 0.1
    addTick pDataSource, lTimestamp2 + 25# * OneMilliSec, TickTypeBid, 0.06
    addTick pDataSource, lTimestamp2 + 40# * OneMilliSec, TickTypeTrade, 0.08
    addTick pDataSource, lTimestamp2 + 55# * OneMilliSec, TickTypeOptionModelPrice, 0.07
    addTick pDataSource, lTimestamp2 + 70# * OneMilliSec, TickTypeOptionModelDelta, 0.1
    addTick pDataSource, lTimestamp2 + 85# * OneMilliSec, -1, 0#
Case "MSFT  200731C00215000"
    addTick pDataSource, lTimestamp2 + 11# * OneMilliSec, TickTypeAsk, 0.02
    addTick pDataSource, lTimestamp2 + 26# * OneMilliSec, TickTypeOptionModelPrice, 0.01
    addTick pDataSource, lTimestamp2 + 41# * OneMilliSec, TickTypeOptionModelDelta, 0#
    addTick pDataSource, lTimestamp2 + 86# * OneMilliSec, -1, 0#
Case Default
    Assert False, "No tick data for contract " & pContractSpec.ToString
End Select

mTaskContext.Continue
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addTick( _
                ByVal pDataSource As MockDataSource, _
                ByVal pTimestamp As Date, _
                ByVal pTickType As TickTypes, _
                ByVal pPrice As Double)
Dim lReplayableTick As ReplayableTick
Set lReplayableTick.DataSource = pDataSource
lReplayableTick.Tick.TimeStamp = pTimestamp
lReplayableTick.Tick.TickType = pTickType
lReplayableTick.Tick.Price = pPrice
mTicks.Add lReplayableTick, pTimestamp
End Sub

Private Function getNextStreamId() As Long
Static sNextStreamId As Long
getNextStreamId = sNextStreamId
sNextStreamId = sNextStreamId + 1
End Function



