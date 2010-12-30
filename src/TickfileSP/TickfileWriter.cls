VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
Implements ITickfileWriter

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "TickfileWriter"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mServiceProviderName As String
Private mTickfileServiceProvider As TickfileServiceProvider

Private mHandle As Long

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileOutputServiceConsumer

Private mTsOut As Scripting.TextStream

Private mTickerID As Long
Private mContract As Contract

Private mPath As String

Private mTickFileVersion As TickFileVersions
Private mTickfileFormat As TickfileFormats

Private mBytesWritten As Long

Private mTickfileGranularity As TickfileGranularities

Private mSessionBuilder As SessionBuilder
Private WithEvents mSession As Session
Attribute mSession.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mPath = App.path & "\TickFiles"
End Sub

Private Sub Class_Terminate()
Debug.Print "TickfileSP.TickfileWriter terminated"
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
mServiceConsumer.Ready
End Sub

'================================================================================
' ITickfileWriter Interface Members
'================================================================================

Private Sub ITickfileWriter_CloseOutputFile()
Const ProcName As String = "ITickfileWriter_CloseOutputFile"
On Error GoTo Err
Finish
mTickfileServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError ProcName
End Sub

Private Sub ITickfileWriter_WriteComment(ByVal pComment As String, ByVal pTimestamp As Date)
mSessionBuilder.SetSessionCurrentTime pTimestamp
writeComment pComment
End Sub

Private Sub ITickfileWriter_WriteTick(tick As TickUtils26.GenericTick)
Const ProcName As String = "ITickfileWriter_WriteTick"
On Error GoTo Err

mSessionBuilder.SetSessionCurrentTime tick.timestamp
writeTickfileLine formatTickData(tick)

Exit Sub
Err:
handleFatalError ProcName
End Sub

'================================================================================
' mSession Event Handlers
'================================================================================

Private Sub mSession_SessionEnded(ev As TimeframeUtils26.SessionEventData)
writeComment "Session ended at " & FormatTimestamp(mSession.CurrentSessionEndTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
End Sub

Private Sub mSession_SessionStarted(ev As TimeframeUtils26.SessionEventData)
Select Case mTickfileGranularity
Case TickfileGranularityDay
    createTickfile mSession.SessionCurrentTime
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
writeComment "Session started at " & FormatTimestamp(mSession.CurrentSessionStartTime, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

closeTickfile
Set mServiceConsumer = Nothing

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pHandle As Long, _
                ByVal Name As String, _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal tickfileOutputServiceConsumer As TradeBuildSP.ITickfileOutputServiceConsumer, _
                ByVal TickfileServiceProvider As TradeBuildSP.ITickfileServiceProvider, _
                ByVal pContract As Contract, _
                ByVal format As TickfileFormats, _
                ByVal version As TickFileVersions, _
                ByVal path As String, _
                ByVal pTickfileGranularity As TickfileGranularities)

mHandle = pHandle
mServiceProviderName = Name
Set mCommonServiceConsumer = CommonServiceConsumer
Set mServiceConsumer = tickfileOutputServiceConsumer
Set mTickfileServiceProvider = TickfileServiceProvider
                
Set mContract = pContract

mTickfileFormat = format
mTickFileVersion = version

If path <> "" Then
    mPath = path
    If Right$(mPath, 1) <> "\" Then mPath = mPath & "\"
End If

mTickfileGranularity = pTickfileGranularity

setupSession

DeferAction Me
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub closeTickfile()
If Not mTsOut Is Nothing Then
    writeComment "Tick recording finished"
    mTsOut.Close
    Set mTsOut = Nothing
    logMessage mServiceProviderName & " (" & mContract.Specifier.LocalSymbol & ")" & vbCrLf & _
                "    Closed tickfile: " & mBytesWritten & "bytes written"
End If
End Sub

Private Sub createCrescendoTickfile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String
Const TICKFILE_CURR_VERSION As Integer = 2

Const ProcName As String = "createCrescendoTickfile"
Dim failpoint As String
On Error GoTo Err

If Not mTsOut Is Nothing Then Exit Sub

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

With mContract.Specifier
    tfPath = tfPath & UCase$(.Symbol) & .Expiry & "\"
    tfName = UCase$(.Symbol) & .Expiry & "-" _
            & format(timestamp, "yyyymmdd-hhmmss") & ".csv"
    Set mTsOut = CreateWriteableTextFile(tfPath & tfName, False)
    
    mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                        "Connected to tickfile " & tfPath & tfName, _
                        mTickfileServiceProvider.handle
    
    mServiceConsumer.OutputTickfileCreated tfPath & tfName
    logMessage "Created output tickfile " & tfPath & tfName
    
    writeTickfileLine "tickfile," & _
                TICKFILE_CURR_VERSION & "," & _
                .Exchange & "," & _
                .Symbol & "," & _
                .Expiry & "," & _
                CDbl(timestamp) & "," & _
                FormatDateTime(timestamp, vbGeneralDate)
End With

Exit Sub
Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub createTickfile(ByVal pTimestamp As Date)
Const ProcName As String = "createTickfile"
Dim failpoint As String
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub createTradeBuildTickFile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String

Const ProcName As String = "createTradeBuildTickFile"
Dim failpoint As String
On Error GoTo Err

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

tfPath = tfPath & UCase$(mContract.Specifier.LocalSymbol) & "\"
tfName = UCase$(mContract.Specifier.LocalSymbol) & "-" & _
        getFilenameDatePart(timestamp) & ".tck"

Set mTsOut = CreateWriteableTextFile(tfPath & tfName, False)

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                    "Connected to tickfile " & tfPath & tfName, _
                    mTickfileServiceProvider.handle

mServiceConsumer.OutputTickfileCreated tfPath & tfName
logMessage "Created output tickfile " & tfPath & tfName

If getFileSize(tfPath & tfName) = 0 Then
    writeTickfileLine TICKFILE_DECLARER & "," & _
                TRADEBUILD_TICKFILE_CURR_VERSION & "," & _
                UCase$(mContract.Specifier.Exchange) & "," & _
                UCase$(mContract.Specifier.Symbol) & "," & _
                mContract.Specifier.Expiry & "," & _
                CDbl(timestamp) & "," & _
                FormatDateTime(timestamp, vbGeneralDate)
    writeTickfileLine CONTRACT_DETAILS_MARKER & mContract.ToXML
Else
    writeComment "Appending to file"
End If
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function formatCrescendoTickData( _
                        ByRef pTick As GenericTick) As String
Dim crescendoTickType As FileTickTypes
Static priceFormatString As String


Const ProcName As String = "formatCrescendoTickData"
Dim failpoint As String
On Error GoTo Err

Select Case pTick.tickType
Case TickTypes.TickTypeAsk
    crescendoTickType = FileTickTypes.ask
Case TickTypes.TickTypeBid
    crescendoTickType = FileTickTypes.bid
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
If crescendoTickType <> 0 Then formatCrescendoTickData = CDbl(pTick.timestamp) & "," & _
                                                        FormatDateTime(pTick.timestamp, vbGeneralDate) & "," & _
                                                        crescendoTickType & "," & _
                                                        pTick.Price & "," & _
                                                        pTick.Size

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function formatTickData( _
                        ByRef pTick As GenericTick) As String
Const ProcName As String = "formatTickData"
Dim failpoint As String
On Error GoTo Err

Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    formatTickData = formatTradeBuildTickData(pTick)
Case TickfileFormats.TickfileCrescendo
    formatTickData = formatCrescendoTickData(pTick)
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function formatTradeBuildTickData( _
                        ByRef pTick As GenericTick) As String
Static priceFormatString As String


Const ProcName As String = "formatTradeBuildTickData"
Dim failpoint As String
On Error GoTo Err

Dim s As String

If priceFormatString = "" Then
    If mContract.NumberOfDecimals = 0 Then
        priceFormatString = "0"
    Else
        priceFormatString = "0." & String(mContract.NumberOfDecimals, "0")
    End If
End If

s = CDbl(pTick.timestamp) & "," & _
    FormatTimestamp(pTick.timestamp, TimestampFormats.TimestampTimeOnly) & ","

Select Case pTick.tickType
Case TickTypes.TickTypeAsk
    s = s & TICK_ASK & "," & format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeBid
    s = s & TICK_BID & "," & format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeClosePrice
    s = s & TICK_CLOSE & "," & format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeHighPrice
    s = s & TICK_HIGH & "," & format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeLowPrice
    s = s & TICK_LOW & "," & format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeMarketDepth
    s = s & TICK_MARKET_DEPTH & "," & _
                pTick.position & "," & _
                pTick.marketmaker & "," & _
                pTick.operation & "," & _
                pTick.side & "," & _
                format(pTick.Price, priceFormatString) & "," & _
                pTick.Size
Case TickTypes.TickTypeMarketDepthReset
    s = s & TICK_MARKET_DEPTH_RESET
Case TickTypes.TickTypeOpenInterest
    s = s & TICK_OPEN_INTEREST & "," & pTick.Size
Case TickTypes.TickTypeOpenPrice
    s = s & TICK_OPEN & "," & format(pTick.Price, priceFormatString)
Case TickTypes.TickTypeTrade
    s = s & TICK_TRADE & "," & format(pTick.Price, priceFormatString) & "," & pTick.Size
Case TickTypes.TickTypeVolume
    s = s & TICK_VOLUME & "," & pTick.Size
End Select

formatTradeBuildTickData = s
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Function getFilenameDatePart(ByVal pTimestamp As Date) As String
Select Case mTickfileGranularity
Case TickfileGranularityDay
    getFilenameDatePart = "Day-" & format(pTimestamp, "yyyymmdd")
Case TickfileGranularityWeek
    getFilenameDatePart = "Week-" & format(WeekStartDate(getSessionDate(pTimestamp)), "yyyymmdd")
Case TickfileGranularitySession
    getFilenameDatePart = "Session-" & format(getSessionDate(pTimestamp), "yyyymmdd")
Case TickfileGranularityExecution
    getFilenameDatePart = "Exec-" & format(pTimestamp, "yyyymmdd-hhmmss")
End Select
End Function

Private Function getFileSize(ByVal pFilename As String) As Long
Dim fso As New FileSystemObject
Dim f As File
Set f = fso.getFile(pFilename)
getFileSize = f.Size
End Function

Private Function getSessionDate(ByVal pTimestamp As Date) As Date
Dim lSessionTimes As SessionTimes
lSessionTimes = mSession.GetSessionTimes(pTimestamp)
If TimeValue(lSessionTimes.startTime) >= CDate("12:00") Then
    getSessionDate = DateValue(lSessionTimes.startTime) + 1
Else
    getSessionDate = DateValue(lSessionTimes.startTime)
End If
End Function

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
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.IsLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message, _
            Me
End Sub

Private Sub setupSession()
Set mSessionBuilder = New SessionBuilder

mSessionBuilder.SessionEndTime = mContract.SessionEndTime
mSessionBuilder.SessionStartTime = mContract.SessionStartTime
Set mSession = mSessionBuilder.Session
End Sub

Private Sub writeComment(ByVal pComment As String)
Const ProcName As String = "writeComment"
On Error GoTo Err

writeTickfileLine "// " & FormatTimestamp(mSession.SessionCurrentTime, TimestampDateAndTimeISO8601) & "  " & pComment

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub writeTickfileLine( _
                ByVal text As String)
Const ProcName As String = "writeTickfileLine"
Dim failpoint As String
On Error GoTo Err

If text = "" Then Exit Sub

mTsOut.WriteLine text
mBytesWritten = mBytesWritten + Len(text)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub




