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
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements DeferredAction
Implements IBarDataReader

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "HistDataReader"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mHistDataServiceProvider As histDataServiceProvider
Private mCommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer
Private mServiceConsumer As TradeBuildSP.IBarDataInputServiceConsumer
Private mDataConsumer As TradeBuildSP.IBarDataConsumer

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1

Private mRequestInProgress As Boolean

Private mRequestId As Long

Private mBarDataSpecifier  As BarDataSpecifier

Private mWaitingForBarData As Boolean
Private mBarRequested As Boolean

Private mSessionBuilder As SessionBuilder
Private mSession As Session

Private mHandle As Long

Private mNumBarsFetched As Long
Private mNumBarsProcessed As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mRequestId = -1
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.HistDataReader terminated"
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Const ProcName As String = "DeferredAction_run"
On Error GoTo Err

mServiceConsumer.Ready
If mRequestId = -1 Then
    ' the request couldn't be serviced
    mServiceConsumer.BarDataComplete
    mRequestInProgress = False
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' IBarDataReader Interface Members
'================================================================================

Private Property Get IBarDataReader_BarDataSpecifier( _
                            ) As BarDataSpecifier
Set IBarDataReader_BarDataSpecifier = mBarDataSpecifier
End Property

Private Sub IBarDataReader_CancelFetch()
Const ProcName As String = "IBarDataReader_CancelFetch"
On Error GoTo Err
If mRequestId <> -1 And Not mTWSAPI Is Nothing Then mTWSAPI.CancelHistoricalData mRequestId
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IBarDataReader_FetchBars( _
                            ByVal barSpecifier As BarDataSpecifier)
Const ProcName As String = "IBarDataReader_FetchBars"


On Error GoTo Err
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
    mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Bar time unit not supported", _
                            mHandle
    Exit Sub
End Select

If barSpecifier.Contract.specifier.sectype = SecTypeCombo Then
    mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                            "Combo contracts not supported", _
                            mHandle
    Exit Sub
End If

mRequestInProgress = True
mWaitingForBarData = True
Set mBarDataSpecifier = barSpecifier

Set mSessionBuilder = New SessionBuilder
mSessionBuilder.SessionStartTime = mBarDataSpecifier.Contract.SessionStartTime
mSessionBuilder.SessionEndTime = mBarDataSpecifier.Contract.SessionEndTime
Set mSession = mSessionBuilder.Session

If mTWSAPI.ConnectionState = ConnConnected Then
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
    DeferAction Me
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IBarDataReader_FireNextBar()
Const ProcName As String = "IBarDataReader_FireNextBar"
On Error GoTo Err
If mWaitingForBarData Then
    mBarRequested = True
Else
    ProcessBar
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Property Get IBarDataReader_NumberOfBars() As Long

End Property

Private Sub IBarDataReader_ReleaseDataStore()
Const ProcName As String = "IBarDataReader_ReleaseDataStore"
On Error GoTo Err

mHistDataServiceProvider.dataReaderFinished Me
Finish

Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Function IBarDataReader_Supports( _
                            ByVal capabilities As Long) As Boolean
IBarDataReader_Supports = gHistDataSupports(capabilities)
End Function

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
Const ProcName As String = "mTWSAPI_Connected"
On Error GoTo Err
If mRequestInProgress Then
    mServiceConsumer.Ready
    mRequestId = mTWSAPI.RequestHistoricalData(mBarDataSpecifier, Me)
End If
Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mHistDataServiceProvider = Nothing
Set mCommonServiceConsumer = Nothing
Set mServiceConsumer = Nothing
Set mTWSAPI = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub HistoricalDataAvailable( _
                ByVal numberOfBars As Long)
Const ProcName As String = "HistoricalDataAvailable"
Dim failpoint As Long
On Error GoTo Err

mWaitingForBarData = False
mNumBarsFetched = numberOfBars
mNumBarsProcessed = 0
If mBarRequested Then
    mBarRequested = False
     mServiceConsumer.Progress 0, 0
    ProcessBar
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal histDataServiceProvider As IBarDataServiceProvider, _
                ByVal CommonServiceConsumer As ICommonServiceConsumer, _
                ByVal histDataInputServiceConsumer As IBarDataInputServiceConsumer, _
                ByVal dataConsumer As IBarDataConsumer, _
                ByVal pTWSAPI As TWSAPI, _
                ByVal Handle As Long)
Set mHistDataServiceProvider = histDataServiceProvider
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = histDataInputServiceConsumer
Set mDataConsumer = dataConsumer
Set mTWSAPI = pTWSAPI
mHandle = Handle
End Sub

Friend Sub HistoricalDataRequestFailed( _
                ByVal reason As String)
Const ProcName As String = "HistoricalDataRequestFailed"
Dim failpoint As Long
On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestFailed, _
                        reason, _
                        mHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub HistoricalDataRequestInvalid( _
                ByVal reason As String)
Const ProcName As String = "HistoricalDataRequestInvalid"
Dim failpoint As Long
On Error GoTo Err

mServiceConsumer.notifyEvent StandardSPEventCodes.HDRequestInvalid, _
                        reason, _
                        mHandle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.number
Dim errSource As String: errSource = Err.source
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
mCommonServiceConsumer.NotifyFatalError Err.number, Err.source, Err.Description, mHandle
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
Dim barTimePeriod As TimePeriod

Dim failpoint As Long
On Error GoTo Err

If mTWSAPI.ReadHistoricalBar(mRequestId, lBar, barTimePeriod) Then
        mDataConsumer.Bar lBar, barTimePeriod
        mNumBarsProcessed = mNumBarsProcessed + 1
        If mNumBarsProcessed Mod 10 = 0 Then mServiceConsumer.Progress mNumBarsProcessed, 100 * mNumBarsProcessed / mNumBarsFetched
        Exit Sub
End If

mServiceConsumer.BarDataComplete
mRequestInProgress = False
mRequestId = -1

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub


