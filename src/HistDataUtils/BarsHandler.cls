VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarsHandler"
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BarsHandler"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mLocalSymbol                                As String

Private mBarType                                    As BarTypes

Private WithEvents mBars                            As Bars
Attribute mBars.VB_VarHelpID = -1
Private mBarWriteController                         As BarWriteController

Private mBarsWritten                                As Long
Private mDatabaseWrites                             As Long

Private mTimePeriod                                 As TimePeriod

Private mBarWriter                                  As IBarWriter

Private mSaveIntervalSeconds                        As Long

Private mStudyBase                                  As IStudyBase
Private mBarStudy                                   As IBarStudy

Private mFutureBuilder                              As New FutureBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mAskBars Event Handlers
'@================================================================================

Private Sub mBars_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mBars_CollectionChanged"
On Error GoTo Err

If ev.ChangeType <> CollItemAdded Then Exit Sub

If Not mBarWriteController Is Nothing Then
    mBarWriteController.Finish
    mDatabaseWrites = mDatabaseWrites + mBarWriteController.DatabaseWrites
End If

Set mBarWriteController = CreateBarWriteController(ev.AffectedItem, mBarType)
mBarsWritten = mBarsWritten + 1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
ElseIf TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    mLocalSymbol = lContract.Specifier.LocalSymbol
ElseIf TypeOf ev.Future.Value Is Bars Then
    mStudyBase.StudyManager.StartStudy mBarStudy, 0
    Set mBars = ev.Future.Value
    gLogger.Log "BarsHandler retrieved " & mBars.Count & " bars for " & mLocalSymbol, ProcName, ModuleName, LogLevelHighDetail
    If mBars.Count <> 0 Then
        Set mBarWriteController = CreateBarWriteController(mBars.Bar(0), mBarType)
        mBarsWritten = mBarsWritten + 1
    End If
    mFutureBuilder.Value = Me
    mFutureBuilder.Complete
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsHandlerFuture() As IFuture
Set BarsHandlerFuture = mFutureBuilder.Future
End Property

Friend Property Get BarsWritten() As Long
BarsWritten = mBarsWritten
End Property

Friend Property Get DatabaseWrites() As Long
DatabaseWrites = mDatabaseWrites
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If mBarWriter Is Nothing Then Exit Sub
If Not mBarWriteController Is Nothing Then
    mBarWriteController.Finish
    mDatabaseWrites = mDatabaseWrites + mBarWriteController.DatabaseWrites
End If

Set mBarWriter = Nothing
Set mBars = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOptions As HistDataWriteOptions, _
                ByVal pBarWriter As IBarWriter, _
                ByVal pSaveIntervalSeconds As Long, _
                ByVal pBarType As BarTypes)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTimePeriod = pTimePeriod
Set mBarWriter = pBarWriter
Set mStudyBase = pStudyBase
mBarType = pBarType

mFutureWaiter.Add pContractFuture

mSaveIntervalSeconds = pSaveIntervalSeconds

Dim lUpdateLatestBar As Boolean
lUpdateLatestBar = pOptions And HistDataWriteOptions.HistDataWriteOptionUpdateLatestBar

Set mBarStudy = createABarStudy(pContractFuture, pClockFuture, pHistDataStore, mBarType, lUpdateLatestBar)
mFutureWaiter.Add mBarStudy.BarsFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createABarStudy( _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarType As BarTypes, _
                ByVal pUpdateLatestBar As Boolean) As IBarStudy
Const ProcName As String = "createABarStudy"
On Error GoTo Err

Dim lBarStudy As IBarStudy

If pUpdateLatestBar Then
    Set lBarStudy = CreateBarStudy(mTimePeriod, mStudyBase, True, fetchLatestBar(pHistDataStore, pContractFuture, pClockFuture, mTimePeriod, pBarType))
Else
    Set lBarStudy = CreateBarStudy(mTimePeriod, mStudyBase, True)
End If
Set createABarStudy = lBarStudy

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function CreateBarWriteController( _
                ByVal pBar As Bar, _
                ByVal pBarType As BarTypes) As BarWriteController
Const ProcName As String = "CreateBarWriteController"
On Error GoTo Err

Dim lBarWriteController As New BarWriteController
lBarWriteController.Initialise pBar, pBarType, mTimePeriod, mLocalSymbol, mSaveIntervalSeconds, mBarWriter
Set CreateBarWriteController = lBarWriteController

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function fetchLatestBar( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pBarType As BarTypes) As IFuture
Const ProcName As String = "fetchLatestBar"
On Error GoTo Err

Dim lBarSpecFuture As IFuture
Set lBarSpecFuture = gCreateBarDataSpecifierFuture(pContractFuture, pTimePeriod, 0, 0, 2, pBarType, pClockFuture, False, True, 0, 0)
Set fetchLatestBar = pHistDataStore.FetchBars(lBarSpecFuture, Nothing)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





