VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataWriter"
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

Implements IBarDataWriter
Implements TaskCompletionListener

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "HistDataWriter"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

Private Type TimeframeEntry
    timeframe               As TimePeriod
    saveIntervalSeconds     As Long
    prevTimestamp           As Date
    barsWritten             As Long
    writes                  As Long
End Type

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                  As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mHistDataServiceProvider    As HistDataServiceProvider

Private mCommonServiceConsumer      As ICommonServiceConsumer
Private mServiceConsumer            As IBarDataOutputServiceConsumer
Private mServiceProviderName        As String

Private mContract                   As Contract

Private mPath                       As String

Private mInstrument                 As Instrument

Private mTimeframeTable(3)          As TimeframeEntry

Private mUseSynchronousWrites       As Boolean

Private WithEvents mInstrumentTC    As TaskController
Attribute mInstrumentTC.VB_VarHelpID = -1

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
Set mTimeframeTable(0).timeframe = GetTimePeriod(1, TimePeriodMinute)
mTimeframeTable(0).saveIntervalSeconds = 5

Set mTimeframeTable(1).timeframe = GetTimePeriod(5, TimePeriodMinute)
mTimeframeTable(1).saveIntervalSeconds = 5

Set mTimeframeTable(2).timeframe = GetTimePeriod(15, TimePeriodMinute)
mTimeframeTable(2).saveIntervalSeconds = 5

Set mTimeframeTable(3).timeframe = GetTimePeriod(60, TimePeriodMinute)
mTimeframeTable(3).saveIntervalSeconds = 5

End Sub

'@===============================================================================
' IBarDataWriter Interface Members
'@===============================================================================

Private Property Get IBarDataWriter_ContractSpecifier() As ContractSpecifier
Set IBarDataWriter_ContractSpecifier = mContract.Specifier
End Property

Private Sub IBarDataWriter_ReleaseDataStore()
Const ProcName As String = "IBarDataWriter_ReleaseDataStore"
On Error GoTo Err
finish
mHistDataServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub IBarDataWriter_WriteBar( _
                ByVal data As Bar, _
                ByVal barType As BarTypes, _
                ByVal barTimePeriod As TimePeriod)
Const ProcName As String = "IBarDataWriter_WriteBar"
Dim index As Long
Dim tc As TaskController

Dim failpoint As Long
On Error GoTo Err

With data
    If mUseSynchronousWrites Then
        mInstrument.WriteBar .TimeStamp, _
                            barType, _
                            barTimePeriod.Length, _
                            .OpenValue, _
                            .HighValue, _
                            .LowValue, _
                            .CloseValue, _
                            .Volume, _
                            .TickVolume, _
                            .OpenInterest
    Else
        Set tc = mInstrument.WriteBarAsync(.TimeStamp, _
                                            barType, _
                                            barTimePeriod.Length, _
                                            .OpenValue, _
                                            .HighValue, _
                                            .LowValue, _
                                            .CloseValue, _
                                            .Volume, _
                                            .TickVolume, _
                                            .OpenInterest)
        tc.AddTaskCompletionListener Me
    End If
    For index = 0 To UBound(mTimeframeTable)
        If mTimeframeTable(index).timeframe.Length = barTimePeriod.Length Then
            With mTimeframeTable(index)
                If .prevTimestamp <> data.TimeStamp Then
                    .barsWritten = .barsWritten + 1
                    .prevTimestamp = data.TimeStamp
                End If
                .writes = .writes + 1
            End With
            Exit For
        End If
    Next
    
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' TaskCompletionListener Interface Members
'@===============================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "TaskCompletionListener_taskCompleted"
Dim failpoint As Long
On Error GoTo Err

If ev.errorNumber <> 0 Then
    logMessage pMsg:="error (" & ev.errorNumber & "): " & ev.errorMessage, pProcName:="TaskCompletionListener_taskCompleted", pLogLevel:=LogLevelHighDetail
    mCommonServiceConsumer.NotifyRecoverableError ev.errorNumber, _
                                        ev.errorMessage, _
                                        mHistDataServiceProvider.handle
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mInstrumentTC Event Handlers
'@===============================================================================

Private Sub mInstrumentTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
Const ProcName As String = "mInstrumentTC_Completed"
Dim failpoint As Long
On Error GoTo Err

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mServiceConsumer.NotifyEvent HDRequestInvalid, _
                                "Error " & ev.errorNumber & ": " & ev.errorMessage, _
                                mHistDataServiceProvider.handle
ElseIf IsEmpty(ev.result) Then
    mServiceConsumer.NotifyEvent HDRequestInvalid, _
                                "Error: specified instrument does not exist: " & mContract.Specifier.ToString, _
                                mHistDataServiceProvider.handle
Else
    Set mInstrument = ev.result
    
    requestTimeframes
    
    mServiceConsumer.Ready
    mServiceConsumer.OutputBarfileCreated mConnectionString
End If
Set mInstrumentTC = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' Properties
'@===============================================================================

Friend Property Let HistDataServiceProvider( _
                ByVal value As HistDataServiceProvider)
Set mHistDataServiceProvider = value
End Property

Friend Property Let Path(ByVal value As String)
mPath = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub finish()
Const ProcName As String = "finish"
Dim s As String
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If Not mInstrumentTC Is Nothing Then
    mInstrumentTC.cancelTask
    Set mInstrumentTC = Nothing
End If

s = mServiceProviderName & " (" & mContract.Specifier.LocalSymbol & ")"
For i = 0 To UBound(mTimeframeTable)
    s = s & vbCrLf & "    Total of " & mTimeframeTable(i).barsWritten & " bars (" & _
            mTimeframeTable(i).timeframe.ToString & _
            ") written in " & _
            mTimeframeTable(i).writes & " database writes"
Next
logMessage pMsg:=s, pProcName:="finish"
Set mServiceConsumer = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub initialise( _
                ByVal Name As String, _
                ByVal dbType As DatabaseTypes, _
                ByVal connectionString As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal barDataOutputServiceConsumer As TradeBuildSP.IBarDataOutputServiceConsumer, _
                ByVal HistDataServiceProvider As TradeBuildSP.IBarDataServiceProvider, _
                ByVal pContract As Contract, _
                ByVal useSynchronousWrites As Boolean)
Const ProcName As String = "initialise"
Dim failpoint As Long
On Error GoTo Err

mServiceProviderName = Name

Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = barDataOutputServiceConsumer

mUseSynchronousWrites = useSynchronousWrites

logMessage pMsg:="Connecting to database", pProcName:="initialise", pLogLevel:=LogLevelDetail

mConnectionString = connectionString
Set mTradingDB = CreateTradingDB(mConnectionString, dbType)

Set mHistDataServiceProvider = HistDataServiceProvider
Set mContract = pContract

If gLogger.IsLoggable(LogLevelDetail) Then logMessage pMsg:="Retrieving instrument object for " & mContract.Specifier.ToString, pProcName:="initialise", pLogLevel:=LogLevelDetail

Set mInstrumentTC = mTradingDB.InstrumentFactory.LoadBySpecifierAsync(mContract.Specifier)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As Long)
Const ProcName As String = "handleFatalError"
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

Private Sub requestTimeframes()
Const ProcName As String = "requestTimeframes"
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

logMessage pMsg:="requesting " & (UBound(mTimeframeTable) + 1) & " timeframes", pProcName:="requestTimeframes", pLogLevel:=LogLevelDetail

For i = 0 To UBound(mTimeframeTable)
    mServiceConsumer.RequiredTimePeriod mTimeframeTable(i).timeframe, _
                                        mTimeframeTable(i).saveIntervalSeconds
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName

End Sub

