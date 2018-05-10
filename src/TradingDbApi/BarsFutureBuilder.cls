VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarsFutureBuilder"
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

Private Const ModuleName                            As String = "BarsFutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1
Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTradingDB                                  As TradingDB

Private mBarSpec                                    As BarDataSpecifier

Private mListener                                   As IBarFetchListener
Private mCookie                                     As Variant

Private mHistDataStore                              As HistDataStore

Private mUseSynchronousReads                        As Boolean

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureBuilder = New FutureBuilder
Set mFutureWaiter = New FutureWaiter
Set mSelfRef = Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureBuilder Event Handlers
'@================================================================================

Private Sub mFutureBuilder_Cancelled(ev As CancelledEventData)
Const ProcName As String = "mFutureBuilder_Cancelled"
On Error GoTo Err

mFutureWaiter.Cancel

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
    mFutureWaiter.Cancel
    mHistDataStore.FetchBuilderFinished Me
    Set mSelfRef = Nothing
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mFutureWaiter.Cancel
    mHistDataStore.FetchBuilderFinished Me
    Set mSelfRef = Nothing
ElseIf TypeOf ev.Future.Value Is TradingDB Then
    Set mTradingDB = ev.Future.Value
    fetchIfReady
ElseIf TypeOf ev.Future.Value Is BarDataSpecifier Then
    Set mBarSpec = ev.Future.Value
    fetchIfReady
ElseIf TypeOf ev.Future.Value Is Bars Then
    mFutureBuilder.Value = ev.Future.Value
    mFutureBuilder.Complete
    mHistDataStore.FetchBuilderFinished Me
    Set mSelfRef = Nothing
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsFuture() As IFuture
Set BarsFuture = mFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pHistDataStore As HistDataStore, _
                ByVal pTradingDBFuture As IFuture, _
                ByVal pBarDataSpecifierFuture As IFuture, _
                ByVal pUseSynchronousReads As Boolean, _
                ByVal pListener As IBarFetchListener, _
                ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

mUseSynchronousReads = pUseSynchronousReads

Set mHistDataStore = pHistDataStore
If pTradingDBFuture.IsAvailable Then
    Set mTradingDB = pTradingDBFuture.Value
Else
    mFutureWaiter.Add pTradingDBFuture
End If

mFutureWaiter.Add pBarDataSpecifierFuture
mFutureBuilder.Cookie = pCookie

Set mListener = pListener
gSetVariant mCookie, pCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fetch()
Const ProcName As String = "fetch"
On Error GoTo Err

If mUseSynchronousReads Then
    mFutureBuilder.Value = mTradingDB.FetchBars( _
                                    mBarSpec.Contract.Specifier, _
                                    mBarSpec.BarTimePeriod, _
                                    mBarSpec.MaxNumberOfBars, _
                                    mBarSpec.FromDate, _
                                    mBarSpec.ToDate, _
                                    mBarSpec.CustomSessionStartTime, _
                                    mBarSpec.CustomSessionEndTime, _
                                    mBarSpec.IncludeBarsOutsideSession, _
                                    mBarSpec.BarType, _
                                    mListener)
    mFutureBuilder.Complete
    mHistDataStore.FetchBuilderFinished Me
Else
    mFutureWaiter.Add mTradingDB.FetchBarsAsync( _
                                    mBarSpec.Contract.Specifier, _
                                    mBarSpec.BarTimePeriod, _
                                    mBarSpec.MaxNumberOfBars, _
                                    mBarSpec.FromDate, _
                                    mBarSpec.ToDate, _
                                    mBarSpec.CustomSessionStartTime, _
                                    mBarSpec.CustomSessionEndTime, _
                                    mBarSpec.IncludeBarsOutsideSession, _
                                    mBarSpec.BarType, _
                                    mCookie, _
                                    mListener)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fetchIfReady()
Const ProcName As String = "fetchIfReady"
On Error GoTo Err

If mTradingDB Is Nothing Or mBarSpec Is Nothing Or mBarSpec.Contract Is Nothing Then Exit Sub

fetch

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


