VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements TradeBuildSP.IBarDataReader

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "HistDataReader"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private WithEvents mBarsTC          As TaskController
Attribute mBarsTC.VB_VarHelpID = -1

Private mHistDataServiceProvider    As HistDataServiceProvider

Private mCommonServiceConsumer      As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer            As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer               As TradeBuildSP.IBarDataConsumer
Private mServiceProviderName        As String

Private mRequestInProgress          As Boolean
            
Private mBarDataSpecifier           As BarDataSpecifier

Private mContract                   As Contract

Private mBars                       As Bars
Private mEnumerator                 As Enumerator

Private mWaitingForData             As Boolean
Private mBarRequested               As Boolean

Private mMode                       As AccessModes

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Terminate()
Debug.Print "HistDataReader terminated"
End Sub

'@===============================================================================
' IBarDataReader Interface Members
'@===============================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As BarDataSpecifier
Set IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
On Error GoTo Err
mRequestInProgress = False
finish
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataReader_CancelFetch", _
                Err.Description
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByVal barSpecifier As BarDataSpecifier)

On Error GoTo Err

gLogger.Log LogLevelDetail, "HistDataReader: fetching bars for " & barSpecifier.Contract.specifier.localSymbol

If mRequestInProgress Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "TBInfoBase.HistDataReader::FetchBars", _
                                    "Request already in progress"

Select Case barSpecifier.barTimePeriod.units
Case TimePeriodSecond, _
    TimePeriodMinute, _
    TimePeriodHour, _
    TimePeriodDay, _
    TimePeriodWeek, _
    TimePeriodMonth, _
    TimePeriodYear, _
    TimePeriodVolume, _
    TimePeriodTickMovement
Case Else
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHistDataServiceProvider.handle
    Exit Sub
End Select

mRequestInProgress = True
mWaitingForData = True

Set mBarDataSpecifier = barSpecifier
Set mContract = mBarDataSpecifier.Contract

Set mBarsTC = mTradingDB.fetchbarsAsync(mBarDataSpecifier.Contract.specifier, _
                                        mBarDataSpecifier.barTimePeriod, _
                                        mBarDataSpecifier.maxNumberOfBars, _
                                        mBarDataSpecifier.FromDate, _
                                        mBarDataSpecifier.ToDate, _
                                        mBarDataSpecifier.customSessionStartTime, _
                                        mBarDataSpecifier.customSessionEndTime, _
                                        mBarDataSpecifier.includeBarsOutsideSession, _
                                        mBarDataSpecifier.barType)

Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataReader_FetchBars", _
                Err.Description

End Sub

Private Sub IBarDataReader_FireNextBar()
On Error GoTo Err
If mWaitingForData Then
    mBarRequested = True
Else
    ProcessBar
End If
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataReader_FireNextBar", _
                Err.Description
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
On Error GoTo Err
finish
mHistDataServiceProvider.dataReaderFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                ProjectName & "." & ModuleName & "::" & "IBarDataReader_ReleaseDataStore", _
                Err.Description
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities, mMode)
End Function

'@===============================================================================
' mBarsTC Event Handlers
'@===============================================================================

Private Sub mBarsTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    gLogger.Log LogLevelNormal, _
                "HistDataReader failed to retrieve bars for " & mBarDataSpecifier.Contract.specifier.toString & vbCrLf & _
                "Error " & ev.errorNumber & ": " & ev.errorMessage
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestFailed, _
                        "Error " & ev.errorNumber & ": " & ev.errorMessage, _
                        mHistDataServiceProvider.handle
Else
    If Not mRequestInProgress Then
        Set mBarsTC = Nothing
        Exit Sub
    End If
    
    mWaitingForData = False
    mRequestInProgress = False
    
    Set mBars = ev.result
    Set mEnumerator = mBars.Enumerator
    
    gLogger.Log LogLevelDetail, "HistDataReader retrieved " & mBars.Count & " bars for " & mBarDataSpecifier.Contract.specifier.toString
    mServiceConsumer.Ready
    
    If mBarRequested Then
        mBarRequested = False
        ProcessBar
    End If
End If
Set mBarsTC = Nothing
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Dim failpoint As Long
On Error GoTo Err

If Not mBarsTC Is Nothing Then
    mBarsTC.cancelTask
    Set mBarsTC = Nothing
End If
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing

Exit Sub

Err:
gLogger.Log LogLevelSevere, _
        "Error " & Err.Number & " at " & _
        ProjectName & "." & ModuleName & ":" & "finish" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
        Err.Description
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal dbType As DatabaseTypes, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataInputServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer, _
                ByVal HistDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal DataConsumer As TradeBuildSP.IBarDataConsumer, _
                ByVal mode As AccessModes)
                
mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataInputServiceConsumer
Set mHistDataServiceProvider = HistDataServiceProvider
Set mDataConsumer = DataConsumer
mMode = mode
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
gLogger.Log LogLevelSevere, Source & ": error (" & Number & "): " & Description

mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mHistDataServiceProvider.handle

finish
mHistDataServiceProvider.dataReaderFinished Me
End Sub

Private Sub ProcessBar()
Dim lBar As Bar

If mEnumerator.moveNext Then
    Set lBar = mEnumerator.current
    mDataConsumer.Bar lBar, mBarDataSpecifier.barTimePeriod
Else
    mServiceConsumer.BarDataComplete
End If
End Sub

