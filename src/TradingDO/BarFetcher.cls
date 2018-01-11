VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarFetcher"
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

Implements IDeferredAction

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

Private Const ModuleName                    As String = "BarFetcher"

Private Const ColumnNameDateTime            As String = "DateTime"
Private Const ColumnNameOpenValue           As String = "OpenPrice"
Private Const ColumnNameHighValue           As String = "HighPrice"
Private Const ColumnNameLowValue            As String = "LowPrice"
Private Const ColumnNameCloseValue          As String = "ClosePrice"
Private Const ColumnNameVolume              As String = "Volume"
Private Const ColumnNameTickVolume          As String = "TickVolume"

Private Const MaxBarsToFetch                As Long = 150000

Private Const TradingDaysPerYear            As Double = 250
Private Const TradingDaysPerMonth           As Double = 21
Private Const TradingDaysPerWeek            As Double = 5

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier                          As BarDataSpecifier

Private mInstrument                         As instrument

Private mDB                                 As TradingDB

Private mCustomSessionBuilder               As SessionBuilder
Private mCustomSession                      As Session

Private WithEvents mTargetBarsBuilder       As BarsBuilder
Attribute mTargetBarsBuilder.VB_VarHelpID = -1
Private mCurrentBar                         As Bar

Private mCookie                             As Variant
Private mBarListener                        As IBarFetchListener

Private mBarsFutureBuilder                  As New FutureBuilder

Private mNotifyAtEnd                        As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "BarFetcher terminated"
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

fireFetchCompleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mTargetBarsBuilder Event Handlers
'@================================================================================

Private Sub mTargetBarsBuilder_BarAdded(ByVal pBar As Bar)
Const ProcName As String = "mTargetBarsBuilder_BarAdded"
On Error GoTo Err

If mNotifyAtEnd Then Exit Sub
fireNotifyBar mCurrentBar

Set mCurrentBar = pBar

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BarsFuture() As IFuture
Set BarsFuture = mBarsFutureBuilder.Future
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchBars( _
                ByVal pSpecifier As BarDataSpecifier, _
                ByVal pListener As IBarFetchListener) As Bars
Const ProcName As String = "FetchBars"
On Error GoTo Err

Set mBarListener = pListener
Set mSpecifier = pSpecifier

Set mTargetBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                            createNewSessionBuilder(mInstrument).Session, _
                                            mInstrument.TickSize, _
                                            IIf(mSpecifier.MaxNumberOfBars = &H7FFFFFFF, 0, mSpecifier.MaxNumberOfBars))

If gUseTickData(mSpecifier.BarTimePeriod.Units) Then
    fetchBarsUsingTickdata
Else
    fetchBarsUsingBardata
End If

fireNotifyBar mCurrentBar

gLogger.Log "BarFetcher fetched " & mTargetBarsBuilder.Bars.Count & " bars for " & mInstrument.ShortName, ProcName, ModuleName

Set FetchBars = mTargetBarsBuilder.Bars
DeferAction Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub Initialise( _
                ByVal db As TradingDB, _
                ByVal pInstrument As instrument)
Set mDB = db
Set mInstrument = pInstrument
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createNewCustomSession()
Const ProcName As String = "createNewCustomSession"
On Error GoTo Err

If mSpecifier.CustomSessionEndTime <> 0 Or mSpecifier.CustomSessionStartTime <> 0 Then
    Set mCustomSessionBuilder = CreateSessionBuilder(mSpecifier.CustomSessionStartTime, mSpecifier.CustomSessionEndTime, GetTimeZone(mInstrument.TimeZoneName))
    Set mCustomSession = mCustomSessionBuilder.Session
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createNewSessionBuilder(ByVal pInstrument As instrument) As SessionBuilder
Const ProcName As String = "createNewSessionBuilder"
On Error GoTo Err

Set createNewSessionBuilder = CreateSessionBuilder(pInstrument.SessionStartTime, pInstrument.SessionEndTime, GetTimeZone(pInstrument.TimeZoneName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchBars( _
                ByVal rsFetcher As RecordsetFetcher, _
                ByVal pInstrumentID As Long, _
                ByRef pReqDetails As BarRequestDetails, _
                ByVal pBarType As Long) As Recordset
Const ProcName As String = "doFetchBars"
On Error GoTo Err

Set doFetchBars = rsFetcher.FetchRecordset( _
                        gSetupFetchBarsCommand( _
                                        pInstrumentID, _
                                        pBarType, _
                                        pReqDetails), _
                        adCmdStoredProc, _
                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function doFetchTicks( _
                ByVal rsFetcher As RecordsetFetcher, _
                ByVal pInstrumentID As Long, _
                ByVal fromTime As Date, _
                ByVal toTime As Date, _
                ByRef pSessionTimes As SessionTimes) As Recordset
Const ProcName As String = "doFetchTicks"
On Error GoTo Err

Set doFetchTicks = rsFetcher.FetchRecordset( _
                                    gSetupFetchTicksCommand(pInstrumentID, _
                                                    fromTime, _
                                                    toTime, _
                                                    pSessionTimes), _
                                    adCmdStoredProc, _
                                        adOpenStatic)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fetchBarsUsingBardata()
Const ProcName As String = "fetchBarsUsingBardata"
On Error GoTo Err

Dim lSessTimes As SessionTimes
lSessTimes = gCalcSessionTimes(mSpecifier, mInstrument)

Dim lReqDetails As BarRequestDetails
lReqDetails = gGenerateBarRequestDetails(lSessTimes, mSpecifier.BarTimePeriod, mSpecifier.FromDate, mSpecifier.ToDate, mSpecifier.MaxNumberOfBars)
If lReqDetails.NotifyAtEnd Then mNotifyAtEnd = True

Dim lRSFetcher As New RecordsetFetcher
lRSFetcher.Initialise mDB

Dim lRS As Recordset
Set lRS = doFetchBars(lRSFetcher, mInstrument.Id, lReqDetails, mSpecifier.BarType)

Dim lBDP As New BarDataProcessor
lBDP.Initialise mTargetBarsBuilder, lRS, lSessTimes, mSpecifier.MaxNumberOfBars, lReqDetails.StartAtFromDate
lBDP.ProcessBars

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub fetchBarsUsingTickdata()
Const ProcName As String = "fetchBarsUsingTickdata"
On Error GoTo Err

Dim lAppending As Boolean
lAppending = (mSpecifier.ToDate = 0)

Dim lSessTimes As SessionTimes
lSessTimes = gCalcSessionTimes(mSpecifier, mInstrument)

Dim lPrevFromDate As Date
Dim lPrevToDate As Date
gGenerateTickRequestDetails lSessTimes, mSpecifier.FromDate, mSpecifier.ToDate, lPrevFromDate, lPrevToDate

Dim lRSFetcher As New RecordsetFetcher
lRSFetcher.Initialise mDB

Dim lTotalInterimBars As Long
Dim lInterimBarsCollection As New EnumerableCollection
Dim lTDP As New TickDataProcessor
Do
    Dim lRS As Recordset
    Set lRS = doFetchTicks(lRSFetcher, mInstrument.Id, lPrevFromDate, lPrevToDate, lSessTimes)
    If lRS.EOF Then Exit Do
    
    Dim lSessionBuilder As SessionBuilder
    Set lSessionBuilder = createNewSessionBuilder(mInstrument)
    
    Dim lInterimBarsBuilder As BarsBuilder
    Set lInterimBarsBuilder = CreateBarsBuilder(mSpecifier.BarTimePeriod, _
                                                lSessionBuilder.Session, _
                                                mInstrument.TickSize)
    lTDP.Initialise lRS, lInterimBarsBuilder, lSessionBuilder, mSpecifier.BarType
    lTDP.ProcessTicks
    
    lInterimBarsCollection.Add lInterimBarsBuilder.Bars
    
    lTotalInterimBars = lTotalInterimBars + lInterimBarsBuilder.Bars.Count
        
    If Not needMoreTicks(lTotalInterimBars, mSpecifier) Then Exit Do
    
    gGenerateNextTickDataRequest lSessTimes, lAppending, lPrevFromDate, lPrevToDate, lPrevFromDate, lPrevToDate
Loop
    
Do While lInterimBarsCollection.Count <> 0
    Dim lBars As Bars
    If lAppending Then
        Set lBars = lInterimBarsCollection(1)
    Else
        Set lBars = lInterimBarsCollection(lInterimBarsCollection.Count)
    End If
    lInterimBarsCollection.Remove lBars
    
    Dim lBar As Bar
    For Each lBar In lBars
        If (lTotalInterimBars <= mTargetBarsBuilder.CacheSize Or _
                mTargetBarsBuilder.CacheSize = 0) And _
            filterByTime(lBar.timestamp) _
        Then mTargetBarsBuilder.AddThisBar lBar
        lTotalInterimBars = lTotalInterimBars - 1
    Next
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function filterByTime( _
                ByVal timestamp As Date) As Boolean
filterByTime = (timestamp >= mSpecifier.FromDate And timestamp < mSpecifier.ToDate)
End Function

Private Sub fireFetchCancelled()
Const ProcName As String = "fireFetchCancelled"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCancelled mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchCompleted()
Const ProcName As String = "fireFetchCompleted"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchCompleted mCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireFetchFailed(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "fireFetchFailed"
On Error GoTo Err

If mBarListener Is Nothing Then Exit Sub

mBarListener.FetchFailed mCookie, pErrorNumber, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireNotifyBar(ByVal pBar As Bar)
Const ProcName As String = "fireNotifyBar"
On Error GoTo Err

If pBar Is Nothing Or mBarListener Is Nothing Then Exit Sub

mBarListener.NotifyBar mCookie, pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function needMoreTicks(ByVal pTotalBars As Long, ByVal pSpecifier As BarDataSpecifier) As Boolean
needMoreTicks = pTotalBars < pSpecifier.MaxNumberOfBars And _
            (pSpecifier.FromDate = 0 Or pSpecifier.ToDate = 0)
End Function

Private Sub notifyBars(ByVal pBars As Bars)
Const ProcName As String = "notifyBars"
On Error GoTo Err

Dim lBar As Bar
For Each lBar In pBars
    fireNotifyBar lBar
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




