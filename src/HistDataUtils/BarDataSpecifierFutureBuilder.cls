VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarDataSpecFutureBldr"
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

Private Const ModuleName                            As String = "BarDataSpecFutureBldr"

'@================================================================================
' Member variables
'@================================================================================

Private mContractFuture                             As IFuture
Private mClockFuture                                As IFuture

Private mBarSpecifier                               As New BarDataSpecifier

Private mExcludeCurrentBar                          As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private WithEvents mFutureBuilder                   As FutureBuilder
Attribute mFutureBuilder.VB_VarHelpID = -1

Private mSelfRef                                    As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
Set mFutureBuilder = New FutureBuilder
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

mFutureWaiter.Clear
mFutureBuilder.Cancel
Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

mBarSpecifier.Contract = mContractFuture.Value
setupSpecifier
mFutureBuilder.Value = mBarSpecifier
mFutureBuilder.Complete
Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    mFutureBuilder.Cancel
    mFutureWaiter.Clear
ElseIf ev.Future.IsFaulted Then
    mFutureBuilder.Fail ev.Future.ErrorNumber, ev.Future.ErrorMessage, ev.Future.ErrorSource
    mFutureWaiter.Clear
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Future() As IFuture
Const ProcName As String = "Future"
On Error GoTo Err

Set Future = mFutureBuilder.Future

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pToTime As Date, _
                ByVal pFromTime As Date, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pClockFuture As IFuture, _
                ByVal pExcludeCurrentBar As Boolean, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pNormaliseDailyTimestamps As Boolean, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractFuture = pContractFuture
Set mClockFuture = pClockFuture

mBarSpecifier.BarTimePeriod = pBarTimePeriod
mBarSpecifier.BarType = pBarType
mBarSpecifier.FromDate = pFromTime
mBarSpecifier.ToDate = pToTime
mBarSpecifier.SetCustomSessionTimes pCustomSessionStartTime, pCustomSessionEndTime
mBarSpecifier.IncludeBarsOutsideSession = pIncludeBarsOutsideSession
mBarSpecifier.MaxNumberOfBars = pMaxNumberOfBars
mBarSpecifier.NormaliseDailyTimestamps = pNormaliseDailyTimestamps
mExcludeCurrentBar = pExcludeCurrentBar

mFutureWaiter.Add mContractFuture
If Not mClockFuture Is Nothing Then mFutureWaiter.Add mClockFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupSpecifier()
Const ProcName As String = "setupSpecifier"
On Error GoTo Err

Dim lContract As IContract
Set lContract = mBarSpecifier.Contract

If mBarSpecifier.ToDate = MaxDate Then
    Dim lClock As Clock
    If Not mClockFuture Is Nothing Then
        Set lClock = mClockFuture.Value
    Else
        Set lClock = GetClock(lContract.TimezoneName)
    End If
    
    Dim lTimestamp As Date
    lTimestamp = lClock.Timestamp
        
    If IsContractExpired(lContract) Then
        Dim lExpiryTimestamp As Date
        lExpiryTimestamp = GetSessionTimes(lContract.ExpiryDate + lContract.SessionEndTime, lContract.SessionStartTime, lContract.SessionEndTime).EndTime
        If lExpiryTimestamp < lTimestamp Then lTimestamp = lExpiryTimestamp
    End If
    
    Dim lSessionStartTime As Date
    lSessionStartTime = IIf(mBarSpecifier.CustomSessionStartTime <> 0, mBarSpecifier.CustomSessionStartTime, mBarSpecifier.Contract.SessionStartTime)
    
    Dim lSessionEndTime As Date
    lSessionEndTime = IIf(mBarSpecifier.CustomSessionEndTime <> 0, mBarSpecifier.CustomSessionEndTime, mBarSpecifier.Contract.SessionEndTime)
    
    If mExcludeCurrentBar And Not IsContractExpired(lContract) Then
        mBarSpecifier.ToDate = BarStartTime(lTimestamp, _
                                    mBarSpecifier.BarTimePeriod, _
                                    lSessionStartTime)
    Else
        ' ensure we get the latest bar
        mBarSpecifier.ToDate = BarEndTime(lTimestamp, _
                                    mBarSpecifier.BarTimePeriod, _
                                    lSessionStartTime, _
                                    lSessionEndTime)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




