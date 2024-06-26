VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timeframe"
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

Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

Event BarsLoaded()

Event BarLoadProgress(ByVal pBarsRetrieved As Long, ByVal pPercentComplete As Single)

Event StateChange(ByRef ev As StateChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                        As String = "Timeframe"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mState                              As TimeframeStates

Private mKey                                As String

Private mStudyBase                          As IStudyBase

Private mInitialBarsFuture                  As IFuture
Private mBarsFuture                         As IFuture
Private mBarStudy                           As IBarStudy

Private mStateChangeListeners               As Listeners

Private mIsHistorical                       As Boolean

Private mContractFuture                     As IFuture
Private mHistDataStore                      As IHistoricalDataStore
Private mFromTime                           As Date
Private mToTime                             As Date
Private mBarType                            As BarTypes
Private mIncludeBarsOutsideSession          As Boolean
Private mExcludeCurrentbar                  As Boolean
Private mClockFuture                        As IFuture

Private WithEvents mFutureWaiter            As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTimePeriod                         As TimePeriod

Private WithEvents mBarFetchListener        As BarFetchListener
Attribute mBarFetchListener.VB_VarHelpID = -1
Private mMaxNumberOfBars                    As Long

Private mCustomSessionEndTime               As Date
Private mCustomSessionStartTime             As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStateChangeListeners = New Listeners
Set mFutureWaiter = New FutureWaiter

' We set the initial state to TimeframeStateFetching because
' that is actually true, since the fetch is initiated as part of the
' construction of the object.
fireStateChange TimeframeStateFetching
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Dim lFetchInitialBars As Boolean: lFetchInitialBars = CBool(Data)

If Not lFetchInitialBars Then
    logAMessage "starting BarStudy", ProcName
    mStudyBase.StudyManager.StartStudy mBarStudy, 0
    fireStateChange TimeframeStateLoaded
    RaiseEvent BarsLoaded
Else
    Set mBarFetchListener = mBarStudy.InitialBarFetchListener
    Dim lBarSpec As BarDataSpecifier
    
    #If SingleDll = 0 Then
    Set lBarSpec = HistDataUtils27.CreateBarDataSpecifier( _
                        mTimePeriod, _
                        mFromTime, _
                        mToTime, _
                        mMaxNumberOfBars, _
                        mBarType, _
                        mExcludeCurrentbar, _
                        mIncludeBarsOutsideSession, _
                        , _
                        mCustomSessionStartTime, _
                        mCustomSessionStartTime)
    #Else
    Set lBarSpec = GHistDataUtils.CreateBarDataSpecifier( _
                        mTimePeriod, _
                        mFromTime, _
                        mToTime, _
                        mMaxNumberOfBars, _
                        mBarType, _
                        mExcludeCurrentbar, _
                        mIncludeBarsOutsideSession, _
                        , _
                        mCustomSessionStartTime, _
                        mCustomSessionStartTime)
    #End If
    Set mInitialBarsFuture = mHistDataStore.FetchBars( _
                                        mContractFuture, _
                                        lBarSpec, _
                                        pClockFuture:=mClockFuture, _
                                        pListener:=mBarFetchListener)
    mFutureWaiter.Add mInitialBarsFuture
End If


Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mBarFetchListener Event Handlers
'@================================================================================

Private Sub mBarFetchListener_FetchStarted(ByVal pCookie As Variant)
Const ProcName As String = "mBarFetchListener_FetchStarted"
On Error GoTo Err

fireStateChange TimeframeStateFetching

Exit Sub

Err:
GTimeframes.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBarFetchListener_FetchDataRetrieved(ByVal pCookie As Variant)
Const ProcName As String = "mBarFetchListener_FetchDataRetrieved"
On Error GoTo Err

fireStateChange TimeframeStateLoading

Exit Sub

Err:
GTimeframes.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
Const ProcName As String = "mBarFetchListener_NotifyBar"
On Error GoTo Err

Static sNumberOfBars As Long
If sNumberOfBars < mMaxNumberOfBars Then sNumberOfBars = sNumberOfBars + 1

Static sETimer As ElapsedTimer
If sETimer Is Nothing Then
    Set sETimer = New ElapsedTimer
    sETimer.StartTiming
End If

If sETimer.ElapsedTimeMicroseconds < 100000 Then Exit Sub

Static sPercent As Single
If mMaxNumberOfBars <> 0 Then
    sPercent = 100! * CSng(sNumberOfBars) / CSng(mMaxNumberOfBars)
Else
    sPercent = sPercent + (100! - sPercent) / 100!
End If

RaiseEvent BarLoadProgress(sNumberOfBars, sPercent)

sETimer.StartTiming

Exit Sub

Err:
GTimeframes.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

logAMessage "starting BarStudy", ProcName
mStudyBase.StudyManager.StartStudy mBarStudy, 0
fireStateChange TimeframeStateLoaded
RaiseEvent BarsLoaded

Exit Sub

Err:
GTimeframes.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub
If ev.Future.IsCancelled Then Exit Sub

If ev.Future.IsFaulted Then
    If ev.Future.ErrorNumber = ErrorCodes.ErrIllegalArgumentException Then
        logAMessage "No bars available: " & ev.Future.ErrorMessage, ProcName
    Else
        Err.Raise ev.Future.ErrorNumber, ev.Future.ErrorSource, ev.Future.ErrorMessage
    End If
Else
    Dim lInitialBars As Bars
    Set lInitialBars = ev.Future.Value
    logAMessage "initial bars loaded: " & lInitialBars.Count, ProcName
End If

Exit Sub

Err:
GTimeframes.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get IsHistorical() As Boolean
IsHistorical = mIsHistorical
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get State() As TimeframeStates
State = mState
End Property

Public Property Get BarsFuture() As IFuture
Set BarsFuture = mBarsFuture
End Property

Public Property Get BarStudy() As IBarStudy
Set BarStudy = mBarStudy
End Property

Public Property Get TimePeriod() As TimePeriod
Set TimePeriod = mTimePeriod
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStateChangeListener( _
                ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mInitialBarsFuture Is Nothing Then
    If mInitialBarsFuture.IsPending Then mInitialBarsFuture.Cancel
    If mBarsFuture.IsPending Then mBarsFuture.Cancel
End If

fireStateChange TimeframeStateFinished

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStudyBase As IStudyBase, _
                ByVal pContractFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pKey As String, _
                ByVal pBarTimePeriod As TimePeriod, _
                ByVal pMaxNumberOfBars As Long, _
                ByVal pBarType As BarTypes, _
                ByVal pFromTime As Date, _
                ByVal pToTime As Date, _
                ByVal pStateChangeListener As IStateChangeListener, _
                ByVal pIncludeBarsOutsideSession As Boolean, _
                ByVal pExcludeCurrentbar As Boolean, _
                ByVal pClockFuture As IFuture, _
                ByVal pCustomSessionStartTime As Date, _
                ByVal pCustomSessionEndTime As Date)
Const ProcName As String = "Initialise"
On Error GoTo Err

If Not pStateChangeListener Is Nothing Then AddStateChangeListener pStateChangeListener

mKey = pKey
Set mStudyBase = pStudyBase
Set mContractFuture = pContractFuture
Set mHistDataStore = pHistDataStore
Set mTimePeriod = pBarTimePeriod
mMaxNumberOfBars = pMaxNumberOfBars
mBarType = pBarType
mFromTime = pFromTime
mToTime = pToTime
mIncludeBarsOutsideSession = pIncludeBarsOutsideSession
mExcludeCurrentbar = pExcludeCurrentbar
Set mClockFuture = pClockFuture
mCustomSessionEndTime = pCustomSessionEndTime
mCustomSessionStartTime = pCustomSessionStartTime

If mToTime <> MaxDate Then mIsHistorical = True

Dim lFetchInitialBars As Boolean
lFetchInitialBars = (mMaxNumberOfBars <> 0 And Not mContractFuture Is Nothing And Not mHistDataStore Is Nothing)
Set mBarStudy = CreateBarStudy(lFetchInitialBars, mTimePeriod, mStudyBase, mIncludeBarsOutsideSession)

Set mBarsFuture = mBarStudy.BarsFuture

DeferAction Me, lFetchInitialBars

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveStateChangeListener( _
                ByVal pListener As IStateChangeListener)
mStateChangeListeners.Remove pListener
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal newState As TimeframeStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

mState = newState

If GTimeframes.Logger.IsLoggable(LogLevelDetail) Then
    Dim s As String
    Select Case newState
    Case TimeframeStateFetching
        s = s & "fetching"
    Case TimeframeStateLoading
        s = s & "loading"
    Case TimeframeStateLoaded
        If mInitialBarsFuture Is Nothing Then
            s = s & "loaded without initial bars"
        Else
            s = s & "loaded with initial bars"
        End If
    Case TimeframeStateFinished
        s = s & "finished"
    End Select
    logAMessage s, ProcName
End If

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = newState

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If
RaiseEvent StateChange(ev)

Exit Sub

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getName() As String
Const ProcName As String = "getName"
On Error GoTo Err

If mContractFuture Is Nothing Then
    getName = "????/" & mKey
ElseIf Not mContractFuture.IsAvailable Then
    getName = "????/" & mKey
Else
    Dim lContract As IContract
    Set lContract = mContractFuture.Value
    getName = lContract.Specifier.LocalSymbol & "/" & mKey
End If

Exit Function

Err:
GTimeframes.HandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logAMessage(ByVal pMessage As String, ByVal pProcName As String, Optional ByVal pLogLevel As LogLevels = LogLevels.LogLevelDetail)
GTimeframes.Logger.Log "Timeframe (" & getName & ") " & pMessage, pProcName, ModuleName, pLogLevel
End Sub




