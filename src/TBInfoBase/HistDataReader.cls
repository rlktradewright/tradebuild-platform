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

Private mNumBarsProcessed           As Long

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
Const ProcName As String = "IBarDataReader_CancelFetch"
On Error GoTo Err
mRequestInProgress = False
finish
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByVal barSpecifier As BarDataSpecifier)
Const ProcName As String = "IBarDataReader_FetchBars"


On Error GoTo Err

logMessage pMsg:="Fetching bars for " & barSpecifier.Contract.Specifier.LocalSymbol, pProcName:="IBarDataReader_FetchBars", pLogLevel:=LogLevelDetail

If mRequestInProgress Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                       ProjectName & "." & ModuleName & ":" & ProcName, _
                                       "Request already in progress"

Select Case barSpecifier.barTimePeriod.Units
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

Set mBarsTC = mTradingDB.FetchbarsAsync(mBarDataSpecifier.Contract.Specifier, _
                                        mBarDataSpecifier.barTimePeriod, _
                                        mBarDataSpecifier.MaxNumberOfBars, _
                                        mBarDataSpecifier.FromDate, _
                                        mBarDataSpecifier.ToDate, _
                                        mBarDataSpecifier.CustomSessionStartTime, _
                                        mBarDataSpecifier.CustomSessionEndTime, _
                                        mBarDataSpecifier.IncludeBarsOutsideSession, _
                                        mBarDataSpecifier.barType)

Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IBarDataReader_FireNextBar()
Const ProcName As String = "IBarDataReader_FireNextBar"
On Error GoTo Err
If mWaitingForData Then
    mBarRequested = True
Else
    ProcessBar
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long
IBarDataReader_NumberOfBars = mBars.Count
End Property

Private Sub IBarDataReader_ReleaseDataStore()
Const ProcName As String = "IBarDataReader_ReleaseDataStore"
On Error GoTo Err
finish
mHistDataServiceProvider.dataReaderFinished Me
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities, mMode)
End Function

'@===============================================================================
' mBarsTC Event Handlers
'@===============================================================================

Private Sub mBarsTC_Completed(ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "mBarsTC_Completed"
Dim failpoint As Long
On Error GoTo Err

If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    logMessage pMsg:="Failed to retrieve bars for " & _
                        mBarDataSpecifier.Contract.Specifier.ToString & vbCrLf & _
                        "Error " & ev.ErrorNumber & ": " & ev.ErrorMessage, _
                pProcName:="mBarsTC_Completed", _
                pLogLevel:=LogLevelSevere
    mServiceConsumer.NotifyEvent StandardSPEventCodes.HDRequestFailed, _
                        "Error " & ev.ErrorNumber & ": " & ev.ErrorMessage, _
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
    
    logMessage pMsg:="Retrieved " & _
                    mBars.Count & " bars for " & mBarDataSpecifier.Contract.Specifier.ToString, _
                pProcName:="mBarsTC_Completed", _
                pLogLevel:=LogLevelDetail
    mServiceConsumer.Ready
    
    mNumBarsProcessed = 0
    mServiceConsumer.Progress 0, 0
    If mBarRequested Then
        mBarRequested = False
        ProcessBar
    End If
End If
Set mBarsTC = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Const ProcName As String = "finish"
Dim failpoint As Long
On Error GoTo Err

If Not mBarsTC Is Nothing Then
    mBarsTC.CancelTask
    Set mBarsTC = Nothing
End If
Set mServiceConsumer = Nothing
Set mDataConsumer = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
Const ProcName As String = "initialise"
                
Dim failpoint As Long
On Error GoTo Err

mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataInputServiceConsumer
Set mHistDataServiceProvider = HistDataServiceProvider
Set mDataConsumer = DataConsumer
mMode = mode
                
mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHistDataServiceProvider.handle
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub ProcessBar()
Const ProcName As String = "ProcessBar"
Dim lBar As Bar

Dim failpoint As Long
On Error GoTo Err

If mEnumerator.MoveNext Then
    Set lBar = mEnumerator.Current
    mDataConsumer.Bar lBar, mBarDataSpecifier.barTimePeriod
    mNumBarsProcessed = mNumBarsProcessed + 1
    If mNumBarsProcessed Mod 10 = 0 Then mServiceConsumer.Progress mNumBarsProcessed, 100 * mNumBarsProcessed / mBars.Count
Else
    mServiceConsumer.BarDataComplete
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

