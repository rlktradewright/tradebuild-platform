VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarStudySupport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Event BarsAvailable(ByVal pBars As Bars)
Event InitialBarsLoaded()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BarStudySupport"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mBarFetchListener                As BarFetchListener
Attribute mBarFetchListener.VB_VarHelpID = -1
Private mInitialBarsReceived                        As Boolean

Private mInitialBarsFutureSupplied                  As Boolean

Private mBarsBuilder                                As BarsBuilder

Private mTimestamp                                  As Date

Private mBarsFutureBuilder                          As FutureBuilder

Private mSession                                    As Session

Private mPriceTickSize                              As Double

Private WithEvents mBars                            As Bars
Attribute mBars.VB_VarHelpID = -1

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mAllowInitialBars                           As Boolean

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
' mBarFetchListener Event Handlers
'@================================================================================

Private Sub mBarFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "mBarFetchListener_FetchCompleted"
On Error GoTo Err

mInitialBarsReceived = True
RaiseEvent InitialBarsLoaded

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
Const ProcName As String = "mBarFetchListener_NotifyBar"
On Error GoTo Err

mBarsBuilder.AddThisBar pBar
mTimestamp = pBar.Timestamp

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsAvailable Then setupBars ev.Future.Value

Exit Sub

Err:
GStudies.NotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BarsBuilder() As BarsBuilder
Set BarsBuilder = mBarsBuilder
End Property

Public Property Get BarsFuture() As IFuture
Const ProcName As String = "BarsFuture"
On Error GoTo Err

Set BarsFuture = mBarsFutureBuilder.Future

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InitialBarFetchListener() As IBarFetchListener
Const ProcName As String = "InitialBarFetchListener"
On Error GoTo Err

AssertArgument mAllowInitialBars, "Initial bars not allowed"
Assert Not mInitialBarsFutureSupplied, "InitialBarsFuture has been supplied"

If mBarFetchListener Is Nothing Then Set mBarFetchListener = New BarFetchListener
Set InitialBarFetchListener = mBarFetchListener

Set mBarsBuilder = GBarUtils.CreateBarsBuilder(mTimePeriod, mSession, mPriceTickSize)
setupBars mBarsBuilder

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let InitialBarsFuture(ByVal Value As IFuture)
Const ProcName As String = "InitialBarsFuture"
On Error GoTo Err

AssertArgument mAllowInitialBars, "Initial bars not allowed"
Assert mBarFetchListener Is Nothing, "InitialBarFetchListener has been requested - InitialBarsFuture cannot be supplied"

mInitialBarsFutureSupplied = True

Set mBarsFutureBuilder = New FutureBuilder
If Value.IsAvailable Then
    GStudies.Logger.Log "Creating BarsBuilder with initial bars", ProcName, ModuleName, LogLevelHighDetail
    setupBars CreateBarsBuilderWithInitialBars(Value.Value, mSession, mPriceTickSize)
Else
    GStudies.Logger.Log "Creating BarsBuilder future with initial bars", ProcName, ModuleName, LogLevelHighDetail
    mFutureWaiter.Add CreateBarsBuilderWithInitialBarsFuture(Value, mSession, mPriceTickSize)
End If

Exit Property

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsReady() As Boolean
IsReady = mBarFetchListener Is Nothing Or mInitialBarsReceived
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSession As Session, _
                ByVal pPriceTickSize As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTimePeriod = pTimePeriod
Set mSession = pSession
mPriceTickSize = pPriceTickSize

Set mBarsFutureBuilder = New FutureBuilder

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Start( _
                ByVal pAllowInitialBars As Boolean)
Const ProcName As String = "Start"
On Error GoTo Err

mAllowInitialBars = pAllowInitialBars
If Not mAllowInitialBars Then
    If mTimePeriod.Length = 0 Then
        setupBars CreateBarsBuilder(mTimePeriod, _
                                    Nothing, _
                                    mPriceTickSize)
    Else
        setupBars CreateBarsBuilder(mTimePeriod, _
                                    mSession, _
                                    mPriceTickSize)
    End If
End If


Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupBars(ByVal pBarsBuilder As BarsBuilder)
Const ProcName As String = "setupBars"
On Error GoTo Err

Set mBarsBuilder = pBarsBuilder
Set mBars = mBarsBuilder.Bars

mBarsFutureBuilder.Value = mBars
mBarsFutureBuilder.Complete

RaiseEvent BarsAvailable(mBars)

Exit Sub

Err:
GStudies.HandleUnexpectedError ProcName, ModuleName
End Sub




