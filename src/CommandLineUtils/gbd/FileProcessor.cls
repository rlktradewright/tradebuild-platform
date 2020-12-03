VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileProcessor"
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

Implements ICollectionChangeListener
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


Private Const ModuleName                            As String = "FileProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mSelfRef                                    As Object

Private mStudyLibraryManager                        As New StudyLibraryManager

Private mTickfileSpec                               As New TickfileSpecifier
Private mTimePeriod                                 As TimePeriod

Private WithEvents mReplayController                As ReplayController
Attribute mReplayController.VB_VarHelpID = -1

Private WithEvents mReplayTaskController            As TaskController
Attribute mReplayTaskController.VB_VarHelpID = -1

Private mNumber                                     As Long
Private mSessionOnly                                As Boolean

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mBars                                       As Bars
Private mBar                                        As Bar

Private mContract                                   As IContract
Private mSecType                                    As SecurityTypes
Private mTickSize                                   As Double

Private mOutputPath                                 As String
Private mOutputFilename                             As String
Private mAppend                                     As Boolean
Private mStream                                     As TextStream

Private mNumberOfBarsOutput                         As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' ICollectionChangeListener Interface Members
'@================================================================================

Private Sub ICollectionChangeListener_Change(ev As CollectionChangeEventData)
Const ProcName As String = "ICollectionChangeListener_Change"
On Error GoTo Err

outputBar mBar
Set mBar = ev.AffectedItem

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

Set IProcessor_ContractSpec = mContract.Specifier

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IProcessor_DataSourceName() As String
IProcessor_DataSourceName = mTickfileSpec.ToString
End Property

Public Property Get IProcessor_FromDate() As Date
IProcessor_FromDate = mTickfileSpec.FromDate
End Property

Private Property Get IProcessor_NumberOfBarsOutput() As Long
IProcessor_NumberOfBarsOutput = mNumberOfBarsOutput
End Property

Private Property Get IProcessor_NumberOfBarsRequested() As Long
IProcessor_NumberOfBarsRequested = mNumber
End Property

Private Property Get IProcessor_Timeframe() As TimePeriod
IProcessor_Timeframe = mTimePeriod
End Property

Public Property Get IProcessor_ToDate() As Date
IProcessor_ToDate = mTickfileSpec.ToDate
End Property

Private Sub IProcessor_StartData( _
                ByVal pOutputPath As String, _
                ByVal pOutputFilename As String, _
                ByVal pAppend As Boolean)
Const ProcName As String = "IProcessor_StartData"
On Error GoTo Err

mOutputPath = pOutputPath
mOutputFilename = pOutputFilename

mStudyLibraryManager.AddBuiltInStudyLibrary

Dim lTickfileSpecs As New TickFileSpecifiers
lTickfileSpecs.Add mTickfileSpec

Dim lTickDataManager As TickfileDataManager
Set lTickDataManager = CreateTickDataManager(lTickfileSpecs, _
                                            CreateTickfileStore(TickfileReadOnly), _
                                            mStudyLibraryManager, _
                                            , _
                                            , _
                                            MarketDataSourceOptNoImpliedTrades + MarketDataSourceOptNoVolumeAdjustments, _
                                            , _
                                            , _
                                            0)
Set mReplayController = lTickDataManager.ReplayController

Dim lTickers As Tickers
Set lTickers = CreateTickers(lTickDataManager, mStudyLibraryManager)

Dim lContractFuture As IFuture
Set lContractFuture = mReplayController.TickStream(0).ContractFuture
mFutureWaiter.Add lContractFuture

Dim lTicker As Ticker
Set lTicker = lTickers.CreateTicker(lContractFuture, False)

Dim lTimeframe As Timeframe
Set lTimeframe = lTicker.Timeframes.Add(mTimePeriod, _
                                        "", _
                                        0, _
                                        0, _
                                        , _
                                        Not mSessionOnly)

If lTimeframe.BarsFuture.IsAvailable Then
    Set mBars = lTimeframe.BarsFuture.Value
    mBars.AddCollectionChangeListener Me
Else
    mFutureWaiter.Add lTimeframe.BarsFuture
End If

lTicker.StartMarketData

Set mReplayTaskController = mReplayController.StartReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IProcessor_StopData()
Const ProcName As String = "IProcessor_StopData"
On Error GoTo Err

mReplayController.StopReplay

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is Bars Then
    Set mBars = ev.Future.Value
    mBars.AddCollectionChangeListener Me
ElseIf TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    mSecType = mContract.Specifier.SecType
    mTickSize = mContract.TickSize

    If mOutputPath = "" And mOutputFilename = "" Then
    Else
        Dim lMessage As String
        Set mStream = gCreateOutputStream( _
                            mOutputPath, _
                            mOutputFilename, _
                            Me, _
                            mAppend, _
                            lMessage)
        If mStream Is Nothing Then
            gNotifyFetchFailed Me, lMessage
            Exit Sub
        End If
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mReplayTaskController Event Handlers
'@================================================================================

Private Sub mReplayTaskController_Completed(ev As TaskCompletionEventData)
Const ProcName As String = "mReplayTaskController_Completed"
On Error GoTo Err

Finish
gNotifyFetchCompleted Me
Set mSelfRef = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTickfileManager Event Handlers
'@================================================================================

Private Sub mReplayController_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "mReplayController_NotifyEvent"
On Error GoTo Err

LogMessage TypeName(ev.Source) & " :" & ev.EventMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mReplayController_ReplayProgress(ByVal pTickfileTimestamp As Date, ByVal pEventsPlayed As Long, ByVal pPercentComplete As Long)
Const ProcName As String = "mReplayController_ReplayProgress"
On Error GoTo Err

gCon.WriteLineToConsole "Processing tickfile: events processed=" & pEventsPlayed & "; percent complete=" & Format(pPercentComplete, "0.0")

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

Friend Sub Initialise( _
                ByVal pTickfileName As String, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date, _
                ByVal pNumber As Long, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pSessionOnly As Boolean, _
                ByVal pEntireSession As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSelfRef = Me

mNumber = pNumber
mSessionOnly = pSessionOnly

mTickfileSpec.FileName = pTickfileName
mTickfileSpec.FromDate = pFromDate
mTickfileSpec.ToDate = pToDate
mTickfileSpec.UseFullSession = Not mSessionOnly
mTickfileSpec.EntireSession = pEntireSession

Set mTimePeriod = pTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

outputBar mBar
mBars.RemoveCollectionChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub outputBar(ByVal pBar As Bar)
Const ProcName As String = "outputBar"
On Error GoTo Err

mNumberOfBarsOutput = mNumberOfBarsOutput + 1
If mStream Is Nothing Then
    gOutputBarToConsole pBar, mSecType, mTickSize
Else
    gOutputBarToTextStream pBar, mSecType, mTickSize, mStream
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

