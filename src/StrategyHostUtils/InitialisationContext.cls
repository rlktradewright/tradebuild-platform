VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InitialisationContext"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "InitialisationContext"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                 As IMarketDataSource
Private mContract                                   As IContract
Private mSession                                    As Session
Private mTimeframes                                 As Timeframes

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mHostController                             As HostController

Private mBarsOutstanding                            As Long

Private mStudyManager                               As StudyManager

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
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    If Not mSession Is Nothing Then mHostController.InitialisationContextIsReady Me
ElseIf TypeOf ev.Future.Value Is Session Then
    Set mSession = ev.Future.Value
    If Not mContract Is Nothing Then mHostController.InitialisationContextIsReady Me
ElseIf TypeOf ev.Future.Value Is Bars Then
    mBarsOutstanding = mBarsOutstanding - 1
    If mBarsOutstanding = 0 Then
        mHostController.InitialisationCompleted
        Set mHostController = Nothing
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Contract() As IContract
Const ProcName As String = "Contract"
On Error GoTo Err

Set Contract = mContract

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsTickReplay() As Boolean
Const ProcName As String = "IsTickReplay"
On Error GoTo Err

IsTickReplay = mDataSource.IsTickReplay

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Session() As Session
Const ProcName As String = "Session"
On Error GoTo Err

Set Session = mSession

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

' Adds a timeframe of the specified bar length to the underlying ticker, and returns
' its name.
'
' @remarks
'   This method can only be called during the <code>IStrategy</code> object&#39;s
'   <code>Initialise</code> method. Calling it elsewhere results in an
'   <code>ErrIllegalStateException</code> error being raised. If a <code>Strategy</code>
'   object needs to create timeframes at other times, it must use the <code>Ticker.Timeframes.Add</code>
'   method.
'
'   Note that the <code>Strategy</code> object&#39;s <code>Start</code> method
'   is not called until all timeframes created using this method have had their
'   historical data fully loaded, and all studies created using the <code>addStudy</code>
'   method have been fully initialised.
' @return
'   An object used to refer to the new timeframe in the <code>TradingContext</code>'s
'   <code>GetBars</code> method.
' @param pTimePeriodLength
'   The time period for this timeframe.
' @param pNumberOfBarsToFetch
'   The number of bars of historical data to be loaded into this timeframe for
'   initialising studies.
' @param pIncludeBarsOutsideSession
'   If set to <code>True</code>, then bars that occur outside the session times
'   defined in the contract for the underlying ticker will be included in the
'   retrieved historical data.
'
'@/
Public Function AddTimeframe( _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pNumberOfBarsToFetch As Long, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean) As Timeframe
Const ProcName As String = "AddTimeframe"
On Error GoTo Err

Dim lTimeframe As Timeframe
Set lTimeframe = mTimeframes.Add(pTimePeriod, "", pNumberOfBarsToFetch, , , pIncludeBarsOutsideSession)
mFutureWaiter.Add lTimeframe.BarsFuture
mBarsOutstanding = mBarsOutstanding + 1

Set AddTimeframe = lTimeframe

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Creates a <code>Study</code> object, taking its input from a specified existing
' <code>Study</code> object.
'
' @remarks
'   This method can only be called during the <code>Strategy</code> object&#39;s
'   <code>Initialise</code> method. Calling it elsewhere results in an
'   <code>ErrIllegalStateException</code> error being raised. If a <code>Strategy</code>
'   object needs to create studies at other times, it must use the
'   <code>Ticker.StudyManager.addStudy</code> method.
'
'   Note that the <code>Strategy</code> object&#39;s <code>Start</code> method
'   is not called until all studies created using this method have been fully initialised,
'   and all timeframes created using the <code>addTimeframe</code> method have had their
'   historical data fully loaded, and .
' @return
'   The new <code>Study</code> object.
' @param pName
'   The name by which the required study is identified in the relevant study library.
' @param underlyingStudy
'   The existing study which is to provide input to the new study.
' @param pInputValueNames
'   An array containing the names of the output values from the underlying study that
'   will be used as input to the new study.
' @param pParams
'   The parameters to be passed to the new study. If this argument is <code>Nothing</code>,
'   the default parameters defined in the study definition will be used.
' @param pLibraryName
'   The name of the study library from which the new <code>Study</code> object is to be
'   created. If this is not supplied or is blank, the new <code>Study</code> object is
'   created from the first study library configured into TradeBuild that can create
'   studies with the required name.
' @param pNumberOfValuesToCache
'   The number of past values that the study should retain. If this argument is omitted, all
'   past values are retained.
'@/
Public Function AddStudy(ByVal pName As String, _
                ByVal pUnderlyingStudy As IStudy, _
                ByRef pInputValueNames() As String, _
                Optional ByVal pParams As Parameters, _
                Optional ByVal pNumberOfValuesToCache As Long, _
                Optional ByVal pLibraryName As String) As IStudy
Const ProcName As String = "AddStudy"
On Error GoTo Err

Set AddStudy = mStudyManager.AddStudy(pName, pUnderlyingStudy, pInputValueNames, pParams, pLibraryName, pNumberOfValuesToCache)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pHostController As HostController, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pStudyManager As StudyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mHostController = pHostController
Set mDataSource = pDataSource
Set mTimeframes = pTimeframes
Set mStudyManager = pStudyManager

mFutureWaiter.Add mDataSource.ContractFuture
mFutureWaiter.Add mDataSource.SessionFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




