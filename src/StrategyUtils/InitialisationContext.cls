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

Implements IStateChangeListener
Implements ITaskCompletionListener

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

Private mStrategyHostRef                            As WeakReference

Private mTicker                                     As Ticker
Private mContract                                   As IContract
Private mSession                                    As Session
Private mClock                                      As Clock
Private mInitialTimestamp                           As Date

Private mPositionManager                            As PositionManager
Private mPositionManagerSimulated                   As PositionManager

Private mStudyManager                               As StudyManager
Private mTimeframes                                 As Timeframes

Private mStrategyRunner                             As StrategyRunner

Private mBarsOutstanding                            As Long
Private mStudiesOutstanding                         As Long

Private mAllowUnprotectedPositions                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim lState As TimeframeStates
lState = ev.State
If lState <> TimeframeStateLoaded Then Exit Sub

mBarsOutstanding = mBarsOutstanding - 1
If mBarsOutstanding = 0 And mStudiesOutstanding = 0 Then
    mStrategyRunner.InitialisationCompleted
    Set mStrategyRunner = Nothing
End If

Dim lTimeframe As Timeframe
Set lTimeframe = ev.Source
lTimeframe.RemoveStateChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ITaskCompletionListener Interface Members
'@================================================================================

Private Sub ITaskCompletionListener_TaskCompleted(ev As TaskCompletionEventData)
Const ProcName As String = "ITaskCompletionListener_TaskCompleted"
On Error GoTo Err

mStudiesOutstanding = mStudiesOutstanding - 1
If mBarsOutstanding = 0 And mStudiesOutstanding = 0 Then
    mStrategyRunner.InitialisationCompleted
    Set mStrategyRunner = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let AllowUnprotectedPositions(ByVal Value As Boolean)
mAllowUnprotectedPositions = Value
End Property

Public Property Get AllowUnprotectedPositions() As Boolean
AllowUnprotectedPositions = mAllowUnprotectedPositions
End Property

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

IsTickReplay = mTicker.IsTickReplay

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositionManager() As PositionManager
Set PositionManager = mPositionManager
End Property

Public Property Get PositionManagerSimulated() As PositionManager
Set PositionManagerSimulated = mPositionManagerSimulated
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

''
' Creates a <code>Study</code> object, taking its input from a specified existing
' <code>Study</code> object.
'
' @remarks
'   This method can only be called during the <code>Strategy</code> object's
'   <code>Initialise</code> method. Calling it elsewhere results in an
'   <code>ErrIllegalStateException</code> error being raised. If a <code>Strategy</code>
'   object needs to create studies at other times, it must use the
'   <code>Ticker.StudyManager.addStudy</code> method.
'
'   Note that the <code>Strategy</code> object's <code>Start</code> method
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
                ByVal pIncludeDataOutsideSession As Boolean, _
                Optional ByVal pParams As Parameters, _
                Optional ByVal pNumberOfValuesToCache As Long, _
                Optional ByVal pLibraryName As String) As IStudy
Const ProcName As String = "AddStudy"
On Error GoTo Err

Set AddStudy = mStudyManager.AddStudy(pName, pUnderlyingStudy, pInputValueNames, pIncludeDataOutsideSession, pParams, pLibraryName, pNumberOfValuesToCache)

Dim lTC As TaskController
Set lTC = mStudyManager.StartStudy(AddStudy, 0)
If Not lTC Is Nothing Then
    lTC.AddTaskCompletionListener Me
    mStudiesOutstanding = mStudiesOutstanding + 1
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' Adds a timeframe of the specified bar length to the underlying ticker, and returns
' its name.
'
' @remarks
'   This method can only be called during the <code>IStrategy</code> object's
'   <code>Initialise</code> method. Calling it elsewhere results in an
'   <code>ErrIllegalStateException</code> error being raised. If a <code>Strategy</code>
'   object needs to create timeframes at other times, it must use the <code>Ticker.Timeframes.Add</code>
'   method.
'
'   Note that the <code>Strategy</code> object's <code>Start</code> method
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
                Optional ByVal pIncludeBarsOutsideSession As Boolean, _
                Optional ByVal pShowInChart As Boolean) As Timeframe
Const ProcName As String = "AddTimeframe"
On Error GoTo Err

Dim lTimeframe As Timeframe
If mTicker.IsTickReplay Then
    Set lTimeframe = mTimeframes.AddHistorical(pTimePeriod, "", pNumberOfBarsToFetch, , BarStartTime(mInitialTimestamp, pTimePeriod, mContract.SessionStartTime), , pIncludeBarsOutsideSession)
Else
    Set lTimeframe = mTimeframes.Add(pTimePeriod, "", pNumberOfBarsToFetch, , , pIncludeBarsOutsideSession)
End If

Static sTimeframes As New EnumerableCollection
If Not sTimeframes.Contains(lTimeframe) Then
    sTimeframes.Add lTimeframe
    lTimeframe.AddStateChangeListener Me
    mBarsOutstanding = mBarsOutstanding + 1
    
    If pShowInChart Then getStrategyHost.AddTimeframe lTimeframe
End If

Set AddTimeframe = lTimeframe

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Displays an output value from a study on the chart.
'
' @remarks
'   This method can only be called during the <code>Strategy</code> object's
'   <code>Initialise</code> method. Calling it elsewhere results in an
'   <code>ErrIllegalStateException</code> error being raised.
'
' @param pStudy
'   The existing study whose value is to be displayed.
' @param pValueName
'   The names of the study value to be displayed on the chart.
' @param pTimeframe
'   Identifies the chart on which the study value is to be displayed.
'   If this is not supplied or is <code>Nothing</code>, the value is displayed
'   on the chart containing the first created timeframe.
'@/
Public Sub ChartStudyValue( _
                ByVal pStudy As IStudy, _
                ByVal pValueName As String, _
                Optional ByVal pTimeframe As Timeframe)
Const ProcName As String = "ChartStudyValue"
On Error GoTo Err

getStrategyHost.ChartStudyValue pStudy, pValueName, pTimeframe


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStrategyHost As IStrategyHost, _
                ByVal pStrategyRunner As StrategyRunner, _
                ByVal pTicker As Ticker)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategyHostRef = CreateWeakReference(pStrategyHost)

Set mStrategyRunner = pStrategyRunner

Set mTicker = pTicker
Set mTimeframes = mTicker.Timeframes
Set mStudyManager = mTicker.StudyBase.StudyManager
Set mPositionManager = mTicker.PositionManager
Set mPositionManagerSimulated = mTicker.PositionManagerSimulated

Set mContract = mTicker.ContractFuture.Value
Set mSession = mTicker.SessionFuture.Value
Set mClock = mTicker.ClockFuture.Value

If mTicker.IsTickReplay Then mInitialTimestamp = mClock.Timestamp

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getStrategyHost() As IStrategyHost
Set getStrategyHost = mStrategyHostRef.Target
End Function



