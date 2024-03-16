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

Private mBarWriter                                  As IBarWriter

Private mSaveIntervalSeconds                        As Long

Private mStudyBase                                  As IStudyBase
Private mBarStudy                                   As IBarStudy

Private mTimePeriod                                 As TimePeriod


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
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If TypeOf ev.Future.Value Is IContract Then
    Dim lContract As IContract
    Set lContract = ev.Future.Value
    mLocalSymbol = lContract.Specifier.LocalSymbol
    GHistData.Logger.Log "Got contract: " & mLocalSymbol, ProcName, ModuleName, LogLevelHighDetail
ElseIf TypeOf ev.Future.Value Is Bars Then
    Set mBars = ev.Future.Value
    GHistData.Logger.Log "BarsHandler retrieved " & mBars.Count & " bars for " & mLocalSymbol, ProcName, ModuleName, LogLevelHighDetail
    mStudyBase.StudyManager.StartStudy mBarStudy, 0
    If mBars.Count <> 0 Then
        Set mBarWriteController = CreateBarWriteController(mBars.Bar(0), mBarType)
        mBarsWritten = mBarsWritten + 1
    End If
End If

Exit Sub

Err:
GHistData.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

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
GHistData.HandleUnexpectedError ProcName, ModuleName
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

GHistData.Logger.Log "Initialising BarsHandler for timeframe " & pTimePeriod.ToShortString, ProcName, ModuleName, LogLevelHighDetail

Set mTimePeriod = pTimePeriod
Set mBarWriter = pBarWriter
Set mStudyBase = pStudyBase
mBarType = pBarType

mFutureWaiter.Add pContractFuture

mSaveIntervalSeconds = pSaveIntervalSeconds

Dim lUpdateLatestBar As Boolean
lUpdateLatestBar = pOptions And HistDataWriteOptions.HistDataWriteOptionUpdateLatestBar

Set mBarStudy = createABarStudy(mTimePeriod, mStudyBase, pContractFuture, pClockFuture, pHistDataStore, mBarType, lUpdateLatestBar)
mFutureWaiter.Add mBarStudy.BarsFuture

Exit Sub

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createABarStudy( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pStudyBase As IStudyBase, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarType As BarTypes, _
                ByVal pUpdateLatestBar As Boolean) As IBarStudy
Const ProcName As String = "createABarStudy"
On Error GoTo Err

GHistData.Logger.Log "Creating bar study for timeperiod: " & pTimePeriod.ToShortString, ProcName, ModuleName, LogLevelHighDetail

Dim lBarStudy As IBarStudy

If pUpdateLatestBar Then
    Set lBarStudy = CreateBarStudy(True, pTimePeriod, pStudyBase, True, fetchLatestBar(pHistDataStore, pContractFuture, pClockFuture, pTimePeriod, pBarType))
Else
    Set lBarStudy = CreateBarStudy(False, pTimePeriod, pStudyBase, True)
End If
Set createABarStudy = lBarStudy

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
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
GHistData.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function fetchLatestBar( _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pClockFuture As IFuture, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pBarType As BarTypes) As IFuture
Const ProcName As String = "fetchLatestBar"
On Error GoTo Err

Dim lBarSpec As BarDataSpecifier
Set lBarSpec = GHistDataUtils.CreateBarDataSpecifier( _
                            pTimePeriod, _
                            0, _
                            0, _
                            2, _
                            pBarType, _
                            False, _
                            True, _
                            False, _
                            0, _
                            0)
Set fetchLatestBar = pHistDataStore.FetchBars( _
                                pContractFuture, _
                                lBarSpec, _
                                "", _
                                "", _
                                pClockFuture, _
                                Nothing)

Exit Function

Err:
GHistData.HandleUnexpectedError ProcName, ModuleName
End Function





