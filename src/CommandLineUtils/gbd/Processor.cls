VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Processor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements IBarFetchListener
Implements ILogListener
Implements IProcessor

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

Private Const ModuleName                            As String = "Processor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mFetchFuture                                As IFuture

Private mStopListening                              As Boolean

Private mContractStore                              As IContractStore
Private mHistDataStore                              As IHistoricalDataStore
Private mContractSpec                               As IContractSpecifier
Private mContract                                   As IContract
Private mSecType                                    As SecurityTypes
Private mTickSize                                   As Double

Private mFromDate                                   As Date
Private mToDate                                     As Date
Private mSpecialFromDate                            As String
Private mSpecialToDate                              As String
Private mNumber                                     As Long
Private mTimePeriod                                 As TimePeriod
Private mSessionOnly                                As Boolean
Private mSessionStartTime                           As Date
Private mSessionEndTime                             As Date
Private mEntireSession                              As Boolean
Private mNormaliseDailyBarTimestamps                As Boolean
Private mIncludeMillisecs                           As Boolean

Private mOutputPath                                 As String
Private mOutputFilename                             As String
Private mAppend                                     As Boolean

Private mNumberOfBarsOutput                         As Long

Private mEchoToStdOut                               As Boolean

Private mFetchStartTime                             As Date
                
'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
If gLogToConsole Then GetLogger("").AddLogListener Me
End Sub

'@================================================================================
' IBarFetchListener Interface Members
'@================================================================================

Private Sub IBarFetchListener_FetchCancelled(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchCancelled"
On Error GoTo Err

If mStopListening Then Exit Sub
mStopListening = True
gNotifyFetchCancelled Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchCompleted(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchCompleted"
On Error GoTo Err

If mStopListening Then Exit Sub
gNotifyFetchCompleted Me, (GetTimestampUTC - mFetchStartTime) * 86400

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchDataRetrieved(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchDataRetrieved"
On Error GoTo Err

If mStopListening Then Exit Sub
gNotifyDataRetrieved Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchFailed(ByVal pCookie As Variant, ByVal pErrorCode As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "IBarFetchListener_FetchFailed"
On Error GoTo Err

If mStopListening Then Exit Sub
gNotifyFetchFailed Me, pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_FetchStarted(ByVal pCookie As Variant)
Const ProcName As String = "IBarFetchListener_FetchStarted"
On Error GoTo Err

If mStopListening Then Exit Sub
mFetchStartTime = GetTimestampUTC
gNotifyFetchStarted Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarFetchListener_NotifyBar(ByVal pCookie As Variant, ByVal pBar As Bar)
Const ProcName As String = "IBarFetchListener_NotifyBar"
On Error GoTo Err

If mStopListening Then Exit Sub
If mFromDate = 0 Then mFromDate = pBar.TimeStamp

Static sStream As TextStream

If mOutputPath = "" And mOutputFilename = "" Then
ElseIf UCase$(mOutputFilename) = "CON" Then
ElseIf sStream Is Nothing Then
    Dim lMessage As String
    Set sStream = gCreateOutputStream( _
                        mOutputPath, _
                        mOutputFilename, _
                        Me, _
                        mAppend, _
                        lMessage)
    If sStream Is Nothing Then
        mFetchFuture.Cancel
        gNotifyFetchFailed Me, lMessage
        mStopListening = True
        Exit Sub
    End If
End If
    
mNumberOfBarsOutput = mNumberOfBarsOutput + 1
If sStream Is Nothing Or mEchoToStdOut Then
    gOutputBarToStdOut mEchoToStdOut, Me, pBar, mTimePeriod, mNumberOfBarsOutput, mNormaliseDailyBarTimestamps, mIncludeMillisecs, mSecType, mTickSize
End If
If Not sStream Is Nothing Then
    gOutputBarToTextStream Me, pBar, mTimePeriod, mNumberOfBarsOutput, mNormaliseDailyBarTimestamps, mIncludeMillisecs, mSecType, mTickSize, sStream
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ILogListener Interface Members
'@================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

Static formatter As ILogFormatter
If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter
gWriteLineToConsole formatter.FormatRecord(Logrec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IProcessor Interface Members
'@================================================================================

Private Property Get IProcessor_ContractSpec() As IContractSpecifier
Const ProcName As String = "IProcessor_ContractSpec"
On Error GoTo Err

If mContract Is Nothing Then
    Set IProcessor_ContractSpec = mContractSpec
Else
    Set IProcessor_ContractSpec = mContract.Specifier
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IProcessor_FromDate() As Date
IProcessor_FromDate = mFromDate
End Property

Private Property Get IProcessor_DataSourceName() As String
Dim s As String
If Not mContract Is Nothing Then
    s = gGetContractName(mContract.Specifier)
Else
    s = mContractSpec.ToString
End If
s = s & " ["
Dim lFrom As String: lFrom = FormatDateTime( _
                                IIf(mEntireSession, Int(mFromDate), mFromDate), _
                                vbGeneralDate)
Dim lTo As String: lTo = IIf(mToDate = MaxDate, _
                                "LATEST", _
                                FormatDateTime( _
                                    IIf(mEntireSession, Int(mToDate), mToDate), _
                                    vbGeneralDate))
If mEntireSession Then s = s & "Sessions "

If mFromDate <> 0 And mToDate <> 0 Then
    s = s & lFrom & " To " & lTo
ElseIf mFromDate <> 0 Then
    s = s & "From " & lFrom
ElseIf mToDate <> 0 Then
    s = s & "To " & lTo
End If
IProcessor_DataSourceName = s & "]"
End Property

Private Property Get IProcessor_NumberOfBarsOutput() As Long
IProcessor_NumberOfBarsOutput = mNumberOfBarsOutput
End Property

Private Property Get IProcessor_NumberOfBarsRequested() As Long
IProcessor_NumberOfBarsRequested = mNumber
End Property

Private Property Get IProcessor_Timeframe() As TimePeriod
Set IProcessor_Timeframe = mTimePeriod
End Property

Public Property Get IProcessor_ToDate() As Date
IProcessor_ToDate = mToDate
End Property

Private Sub IProcessor_StartData( _
                ByVal pOutputPath As String, _
                ByVal pOutputFilename As String, _
                ByVal pAppend As Boolean)
Const ProcName As String = "IProcessor_StartData"
On Error GoTo Err

mOutputPath = pOutputPath
mOutputFilename = pOutputFilename
mAppend = pAppend

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(mContractSpec, mContractStore)
mFutureWaiter.Add lContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IProcessor_StopData()
Const ProcName As String = "IProcessor_StopData"
On Error GoTo Err

mFutureWaiter.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitAllCompleted(ev As TWUtilities40.FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitAllCompleted"
On Error GoTo Err

Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsCancelled Then
    gNotifyFetchCancelled Me
    mStopListening = True
ElseIf ev.Future.IsFaulted Then
    gNotifyFetchFailed Me, ev.Future.ErrorMessage
    mStopListening = True
ElseIf ev.Future.Value Is Nothing Then
    ' bars fetch has completed
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    mSecType = mContract.Specifier.SecType
    mTickSize = mContract.TickSize
    
    If mEntireSession Then
        If mFromDate <> 0 Then
            mFromDate = GetSessionTimes( _
                                Int(mFromDate) + 0.5, _
                                IIf(mSessionOnly, mContract.SessionStartTime, mContract.FullSessionStartTime), _
                                IIf(mSessionOnly, mContract.SessionEndTime, mContract.FullSessionEndTime)).StartTime
        End If
        If mToDate <> 0 And mToDate <> MaxDate Then
            mToDate = GetSessionTimes( _
                                Int(mToDate) + 0.5, _
                                IIf(mSessionOnly, mContract.SessionStartTime, mContract.FullSessionStartTime), _
                                IIf(mSessionOnly, mContract.SessionEndTime, mContract.FullSessionEndTime)).EndTime
        End If
    End If
    
    Dim lBarDataSpec As BarDataSpecifier
    Set lBarDataSpec = CreateBarDataSpecifier( _
                                mTimePeriod, _
                                mFromDate, _
                                mToDate, _
                                mNumber, _
                                BarTypes.BarTypeTrade, _
                                False, _
                                Not mSessionOnly, _
                                mNormaliseDailyBarTimestamps, _
                                mSessionStartTime, _
                                mSessionEndTime)
    
    ' Note no Bars object will be returned
    Set mFetchFuture = mHistDataStore.FetchBars( _
                            ev.Future, _
                            lBarDataSpec, _
                            mSpecialFromDate, _
                            mSpecialToDate, _
                            pListener:=Me, _
                            pDontReturnBars:=True)
                            
    mFutureWaiter.Add mFetchFuture
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

GetLogger("").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date, _
                ByVal pSpecialFromDate As String, _
                ByVal pSpecialToDate As String, _
                ByVal pNumber As Long, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionOnly As Boolean, _
                ByVal pSessionStartTime As Date, _
                ByVal pSessionEndTime As Date, _
                ByVal pEntireSession As Boolean, _
                ByVal pNormaliseDailyBarTimestamps As Boolean, _
                ByVal pIncludeMillisecs As Boolean, _
                ByVal pEchoToStdOut As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mContractStore = pContractStore
Set mHistDataStore = pHistDataStore
Set mContractSpec = pContractSpec
mFromDate = pFromDate
mSpecialFromDate = pSpecialFromDate
mSpecialToDate = pSpecialToDate
mToDate = pToDate
mNumber = pNumber
Set mTimePeriod = pTimePeriod
mSessionOnly = pSessionOnly
mSessionStartTime = pSessionStartTime
mSessionEndTime = pSessionEndTime
mEntireSession = pEntireSession
mNormaliseDailyBarTimestamps = pNormaliseDailyBarTimestamps
mIncludeMillisecs = pIncludeMillisecs
mEchoToStdOut = pEchoToStdOut

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

