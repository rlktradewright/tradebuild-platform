VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextTickfileWriter"
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

Implements ITickfileWriter

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TextTickfileWriter"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickfileStore As TickfileStore

Private mOutputMonitor As ITickfileOutputMonitor

Private mTsOut As Scripting.TextStream

Private mContractFuture As IFuture
Private mContract As IContract

Private mPath As String

Private mTickFileVersion As TickFileVersions
Private mTickfileFormat As TickfileFormats

Private mBytesWritten As Long

Private mTickfileGranularity As TickfileGranularities

Private mSessionBuilder As SessionBuilder
Private WithEvents mSession As Session
Attribute mSession.VB_VarHelpID = -1

Private mReady As Boolean

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mPath = ApplicationSettingsFolder & "\TickFiles"
Set mFutureWaiter = New FutureWaiter
End Sub

'================================================================================
' ITickfileWriter Interface Members
'================================================================================

Private Property Get ITickfileWriter_ContractFuture() As IFuture
Set ITickfileWriter_ContractFuture = mContractFuture
End Property

Private Sub ITickfileWriter_Finish()
Const ProcName As String = "ITickfileWriter_Finish"
On Error GoTo Err

Assert mReady, "Not ready"

Finish
mTickfileStore.DataWriterFinished Me

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ITickfileWriter_IsReady() As Boolean
ITickfileWriter_IsReady = mReady
End Property

Private Sub ITickfileWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
Const ProcName As String = "ITickfileWriter_WriteComment"
On Error GoTo Err

Assert mReady, "Not ready"

mSessionBuilder.SetSessionCurrentTime pTimestamp
WriteComment pComment

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileWriter_WriteTick(Tick As GenericTick)
Const ProcName As String = "ITickfileWriter_WriteTick"
On Error GoTo Err

Assert mReady, "Not ready"

mSessionBuilder.SetSessionCurrentTime Tick.Timestamp
writeTickfileLine formatTickData(Tick)

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

If TypeOf ev.Future.Value Is IContract Then
    Set mContract = ev.Future.Value
    setupSession
    setReady
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionEnded(ev As SessionEventData)
Const ProcName As String = "mSession_SessionEnded"
On Error GoTo Err

If mTsOut Is Nothing Then createTickfile mSession.SessionCurrentTime

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mSession_SessionStarted(ev As SessionEventData)
Const ProcName As String = "mSession_SessionStarted"
On Error GoTo Err

Select Case mTickfileGranularity
Case TickfileGranularityDay
    If mTsOut Is Nothing Then createTickfile mSession.SessionCurrentTime
Case TickfileGranularityWeek
    If mTsOut Is Nothing Or _
        mSession.IsTimeInFirstSessionForWeek(mSession.SessionCurrentTime) _
    Then
        createTickfile mSession.SessionCurrentTime
    End If
Case TickfileGranularitySession
    createTickfile mSession.SessionCurrentTime
Case TickfileGranularityExecution
    If mTsOut Is Nothing Then createTickfile mSession.SessionCurrentTime
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

closeTickfile
mOutputMonitor.NotifyOutputFileClosed Me
Set mOutputMonitor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pTickfileStore As TickfileStore, _
                ByVal pContractFuture As IFuture, _
                ByVal pFormat As TickfileFormats, _
                ByVal pVersion As TickFileVersions, _
                ByVal pPath As String, _
                ByVal pTickfileGranularity As TickfileGranularities)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOutputMonitor = pOutputMonitor
Set mTickfileStore = pTickfileStore
                
Set mContractFuture = pContractFuture

mTickfileFormat = pFormat
mTickFileVersion = pVersion

If pPath <> "" Then
    mPath = pPath
    If Right$(mPath, 1) <> "\" Then mPath = mPath & "\"
End If

mTickfileGranularity = pTickfileGranularity

mFutureWaiter.Add pContractFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub closeTickfile()
Const ProcName As String = "closeTickfile"
On Error GoTo Err

If Not mTsOut Is Nothing Then
    mTsOut.Close
    Set mTsOut = Nothing
    gLogger.Log "Closed tickfile (" & mContract.Specifier.LocalSymbol & "): " & vbCrLf & _
                mBytesWritten & " bytes written", _
                ProcName, ModuleName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createCrescendoTickfile(ByVal Timestamp As Date)
Const ProcName As String = "createCrescendoTickfile"
On Error GoTo Err

Dim tfPath As String
Dim tfName As String
Const TICKFILE_CURR_VERSION As Integer = 2

If Not mTsOut Is Nothing Then Exit Sub

tfPath = IIf(mPath <> "", mPath, App.Path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

With mContract.Specifier
    tfPath = tfPath & UCase$(.Symbol) & .Expiry & "\"
    tfName = UCase$(.Symbol) & .Expiry & "-" _
            & Format(Timestamp, "yyyymmdd-hhmmss") & ".csv"
    Set mTsOut = CreateWriteableTextFile(tfPath & tfName, False)
    
    mOutputMonitor.NotifyOutputFileCreated Me, tfPath & tfName
    gLogger.Log "Created output tickfile " & tfPath & tfName, ProcName, ModuleName
    
    writeTickfileLine "tickfile," & _
                TICKFILE_CURR_VERSION & "," & _
                .Exchange & "," & _
                .Symbol & "," & _
                .Expiry & "," & _
                CDbl(Timestamp) & "," & _
                FormatDateTime(Timestamp, vbGeneralDate)
End With

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createTickfile(ByVal pTimestamp As Date)
Const ProcName As String = "createTickfile"
On Error GoTo Err

Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    closeTickfile
    createTradeBuildTickFile pTimestamp
Case TickfileFormats.TickfileCrescendo
    createCrescendoTickfile pTimestamp
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createTradeBuildTickFile(ByVal Timestamp As Date)
Const ProcName As String = "createTradeBuildTickFile"
On Error GoTo Err

Dim tfPath As String
Dim tfName As String

tfPath = IIf(mPath <> "", mPath, App.Path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

tfPath = tfPath & UCase$(mContract.Specifier.LocalSymbol) & "\"
tfName = UCase$(mContract.Specifier.LocalSymbol) & "-" & _
        getFilenameDatePart(Timestamp) & ".tck"

Set mTsOut = CreateWriteableTextFile(tfPath & tfName, False)

mOutputMonitor.NotifyOutputFileCreated Me, tfPath & tfName
gLogger.Log "Created output tickfile " & tfPath & tfName, ProcName, ModuleName

If getFileSize(tfPath & tfName) = 0 Then
    writeTickfileLine TICKFILE_DECLARER & "," & _
                TRADEBUILD_TICKFILE_CURR_VERSION & "," & _
                UCase$(mContract.Specifier.Exchange) & "," & _
                UCase$(mContract.Specifier.Symbol) & "," & _
                mContract.Specifier.Expiry & "," & _
                CDbl(Timestamp) & "," & _
                FormatDateTime(Timestamp, vbGeneralDate)
    writeTickfileLine CONTRACT_DETAILS_MARKER & ContractToXML(mContract)
Else
    WriteComment "Appending to file"
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function formatCrescendoTickData( _
                        ByRef pTick As GenericTick) As String
Const ProcName As String = "formatCrescendoTickData"
On Error GoTo Err

Dim crescendoTickType As FileTickTypes
Static priceFormatString As String

Select Case pTick.TickType
Case TickTypes.TickTypeAsk
    crescendoTickType = FileTickTypes.Ask
Case TickTypes.TickTypeBid
    crescendoTickType = FileTickTypes.Bid
Case TickTypes.TickTypeClosePrice
    crescendoTickType = FileTickTypes.PrevClose
Case TickTypes.TickTypeHighPrice
    crescendoTickType = FileTickTypes.High
Case TickTypes.TickTypeLowPrice
    crescendoTickType = FileTickTypes.Low
Case TickTypes.TickTypeTrade
    crescendoTickType = FileTickTypes.Last
Case TickTypes.TickTypeVolume
    crescendoTickType = FileTickTypes.Volume
End Select
If crescendoTickType <> 0 Then formatCrescendoTickData = CDbl(pTick.Timestamp) & "," & _
                                                        FormatDateTime(pTick.Timestamp, vbGeneralDate) & "," & _
                                                        crescendoTickType & "," & _
                                                        pTick.Price & "," & _
                                                        pTick.Size

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatTickData( _
                        ByRef pTick As GenericTick) As String
Const ProcName As String = "formatTickData"
On Error GoTo Err

Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    formatTickData = formatTradeBuildTickData(pTick)
Case TickfileFormats.TickfileCrescendo
    formatTickData = formatCrescendoTickData(pTick)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function formatTradeBuildTickData( _
                        ByRef pTick As GenericTick) As String
Const ProcName As String = "formatTradeBuildTickData"
On Error GoTo Err

Static priceFormatString As String
Dim s As String

If priceFormatString = "" Then
    If mContract.NumberOfDecimals = 0 Then
        priceFormatString = "0"
    Else
        priceFormatString = "0." & String(mContract.NumberOfDecimals, "0")
    End If
End If

s = CDbl(pTick.Timestamp) & "," & _
    FormatTimestamp(pTick.Timestamp, TimestampFormats.TimestampTimeOnly) & ","

Select Case pTick.TickType
Case TickTypes.TickTypeAsk
    s = s & TICK_ASK & "," & Format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeBid
    s = s & TICK_BID & "," & Format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeClosePrice
    s = s & TICK_CLOSE & "," & Format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeHighPrice
    s = s & TICK_HIGH & "," & Format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeLowPrice
    s = s & TICK_LOW & "," & Format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeMarketDepth
    s = s & TICK_MARKET_DEPTH & "," & _
                pTick.Position & "," & _
                pTick.MarketMaker & "," & _
                pTick.Operation & "," & _
                pTick.Side & "," & _
                Format(pTick.Price, priceFormatString) & "," & _
                pTick.Size
Case TickTypes.TickTypeMarketDepthReset
    s = s & TICK_MARKET_DEPTH_RESET
Case TickTypes.TickTypeOpenInterest
    s = s & TICK_OPEN_INTEREST & "," & pTick.Size
Case TickTypes.TickTypeOpenPrice
    s = s & TICK_OPEN & "," & Format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeTrade
    s = s & TICK_TRADE & "," & Format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeVolume
    s = s & TICK_VOLUME & "," & pTick.Size
End Select

formatTradeBuildTickData = s
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFilenameDatePart(ByVal pTimestamp As Date) As String
Const ProcName As String = "getFilenameDatePart"
On Error GoTo Err

Select Case mTickfileGranularity
Case TickfileGranularityDay
    getFilenameDatePart = "Day-" & Format(pTimestamp, "yyyymmdd")
Case TickfileGranularityWeek
    getFilenameDatePart = "Week-" & Format(WeekStartDate(getSessionDate(pTimestamp)), "yyyymmdd")
Case TickfileGranularitySession
    getFilenameDatePart = "Session-" & Format(getSessionDate(pTimestamp), "yyyymmdd")
Case TickfileGranularityExecution
    getFilenameDatePart = "Exec-" & Format(pTimestamp, "yyyymmdd-hhmmss")
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFileSize(ByVal pFilename As String) As Long
Const ProcName As String = "getFileSize"
On Error GoTo Err

Dim fso As New FileSystemObject
Dim f As File
Set f = fso.GetFile(pFilename)
getFileSize = f.Size

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSessionDate(ByVal pTimestamp As Date) As Date
Const ProcName As String = "getSessionDate"
On Error GoTo Err

Dim lSessionTimes As SessionTimes

lSessionTimes = mSession.GetSessionTimes(pTimestamp)
If TimeValue(lSessionTimes.StartTime) >= CDate("12:00") Then
    getSessionDate = DateValue(lSessionTimes.StartTime) + 1
Else
    getSessionDate = DateValue(lSessionTimes.StartTime)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setReady()
Const ProcName As String = "setReady"
On Error GoTo Err

mReady = True
If Not mOutputMonitor Is Nothing Then mOutputMonitor.NotifyReady Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSession()
Const ProcName As String = "setupSession"
On Error GoTo Err

Set mSessionBuilder = CreateSessionBuilder(mContract.sessionStartTime, mContract.sessionEndTime, Nothing)
Set mSession = mSessionBuilder.Session

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub WriteComment(ByVal pComment As String)
Const ProcName As String = "writeComment"
On Error GoTo Err

writeTickfileLine "// " & FormatTimestamp(mSession.SessionCurrentTime, TimestampDateAndTimeISO8601) & "  " & pComment

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writeTickfileLine( _
                ByVal text As String)
Const ProcName As String = "writeTickfileLine"
On Error GoTo Err

If text = "" Then Exit Sub
If mTsOut Is Nothing Then Exit Sub

mTsOut.WriteLine text
mBytesWritten = mBytesWritten + Len(text)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




