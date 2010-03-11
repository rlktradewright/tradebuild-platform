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

Private mFileSys As Scripting.FileSystemObject
Private mTsOut As Scripting.TextStream

Private mTickerID As Long
Private mContract As Contract

Private mPath As String

Private mTickFileVersion As TickFileVersions
Private mTickfileFormat As TickfileFormats

Private mBytesWritten As Long

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mPath = App.path & "\TickFiles"
Set mFileSys = New FileSystemObject
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

Private Sub ITickfileWriter_WriteTick(tick As TickUtils26.GenericTick)
Const ProcName As String = "ITickfileWriter_WriteTick"
On Error GoTo Err

Select Case tick.tickType
Case TickTypes.TickTypeAsk
    writeTickData tick.timestamp, TICK_ASK, tick.price, tick.size
Case TickTypes.TickTypeBid
    writeTickData tick.timestamp, TICK_BID, tick.price, tick.size
Case TickTypes.TickTypeClosePrice
    writeTickData tick.timestamp, TICK_CLOSE, tick.price
Case TickTypes.TickTypeHighPrice
    writeTickData tick.timestamp, TICK_HIGH, tick.price
Case TickTypes.TickTypeLowPrice
    writeTickData tick.timestamp, TICK_LOW, tick.price
Case TickTypes.TickTypeMarketDepth
    Select Case mTickfileFormat
    Case TickfileFormats.TickfileTradeBuild
        If mTsOut Is Nothing Then createTradeBuildTickFile tick.timestamp
        
        writeTickfileLine CDbl(tick.timestamp) & "," & _
                    FormatTimestamp(tick.timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                    TICK_MARKET_DEPTH & "," & _
                    tick.position & "," & _
                    tick.marketmaker & "," & _
                    tick.operation & "," & _
                    tick.side & "," & _
                    format(tick.price, "Fixed") & "," & _
                    tick.size
    Case TickfileFormats.TickfileCrescendo
        ' Crescendo can't handle market depth in tickfiles
    End Select
Case TickTypes.TickTypeMarketDepthReset
    Select Case mTickfileFormat
    Case TickfileFormats.TickfileTradeBuild
        If mTsOut Is Nothing Then createTradeBuildTickFile tick.timestamp
        
        writeTickfileLine CDbl(tick.timestamp) & "," & _
                    FormatTimestamp(tick.timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                    TICK_MARKET_DEPTH_RESET
    
    Case TickfileFormats.TickfileCrescendo
        ' Crescendo can't handle market depth in tickfiles
    End Select
Case TickTypes.TickTypeOpenInterest
    writeTickData tick.timestamp, TICK_OPEN_INTEREST, , tick.size
Case TickTypes.TickTypeOpenPrice
    writeTickData tick.timestamp, TICK_OPEN, tick.price
Case TickTypes.TickTypeTrade
    writeTickData tick.timestamp, TICK_TRADE, tick.price, tick.size
Case TickTypes.TickTypeVolume
    writeTickData tick.timestamp, TICK_VOLUME, , tick.size
End Select

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
Dim failpoint As String
On Error GoTo Err

If Not mTsOut Is Nothing Then
    mTsOut.Close
    Set mTsOut = Nothing
    logMessage mServiceProviderName & " (" & mContract.Specifier.LocalSymbol & ")" & vbCrLf & _
                "    Closed tickfile: " & mBytesWritten & "bytes written"
End If
Set mServiceConsumer = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
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
                ByVal path As String)

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

DeferAction Me
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub createCrescendoTickfile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String
Const TICKFILE_CURR_VERSION As Integer = 2

Const ProcName As String = "createCrescendoTickfile"
Dim failpoint As String
On Error GoTo Err

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

On Error Resume Next
mFileSys.CreateFolder tfPath
On Error GoTo Err

With mContract.Specifier
    tfPath = tfPath & UCase$(.Symbol) & .Expiry & "\"
    tfName = UCase$(.Symbol) & .Expiry & "-" _
            & format(timestamp, "yyyymmdd-hhmmss") & ".csv"
    Set mTsOut = mFileSys.CreateTextFile(tfPath & tfName)
    
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
If Err.Number = 76 Then
    mFileSys.CreateFolder tfPath
    Resume
End If
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub createTradeBuildTickFile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String

Const ProcName As String = "createTradeBuildTickFile"
Dim failpoint As String
On Error GoTo Err

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

On Error Resume Next
mFileSys.CreateFolder tfPath
On Error GoTo Err

tfPath = tfPath & UCase$(mContract.Specifier.LocalSymbol) & "\"
tfName = UCase$(mContract.Specifier.LocalSymbol) & "-" & _
        format(timestamp, "yyyymmdd-hhmmss") & ".tck"
Set mTsOut = mFileSys.CreateTextFile(tfPath & tfName)

mServiceConsumer.NotifyEvent StandardSPEventCodes.TFConnectedToDataSource, _
                    "Connected to tickfile " & tfPath & tfName, _
                    mTickfileServiceProvider.handle

mServiceConsumer.OutputTickfileCreated tfPath & tfName
logMessage "Created output tickfile " & tfPath & tfName
writeTickfileLine TICKFILE_DECLARER & "," & _
            TRADEBUILD_TICKFILE_CURR_VERSION & "," & _
            UCase$(mContract.Specifier.Exchange) & "," & _
            UCase$(mContract.Specifier.Symbol) & "," & _
            mContract.Specifier.Expiry & "," & _
            CDbl(timestamp) & "," & _
            FormatDateTime(timestamp, vbGeneralDate)
writeTickfileLine CONTRACT_DETAILS_MARKER & mContract.ToXML

Exit Sub

Err:
If Err.Number = 76 Then
    mFileSys.CreateFolder tfPath
    Resume
End If
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

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

Private Sub writeTickData( _
                        ByVal timestamp As Date, _
                        ByVal tickType As String, _
                        Optional ByVal price As Double = -1#, _
                        Optional ByVal size As Long = -1&)
Dim crescendoTickType As FileTickTypes
Static priceFormatString As String


Const ProcName As String = "writeTickData"
Dim failpoint As String
On Error GoTo Err

If priceFormatString = "" Then
    If mContract.NumberOfDecimals = 0 Then
        priceFormatString = "0"
    Else
        priceFormatString = "0." & String(mContract.NumberOfDecimals, "0")
    End If
End If

Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    If mTsOut Is Nothing Then createTradeBuildTickFile timestamp
    writeTickfileLine CDbl(timestamp) & "," & _
                FormatTimestamp(timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                tickType & "," & _
                IIf(price <> -1#, format(price, priceFormatString), "") & _
                IIf(size = -1&, "", IIf(price = -1#, size, "," & size))
Case TickfileFormats.TickfileCrescendo
    If mTsOut Is Nothing Then createCrescendoTickfile timestamp
    Select Case tickType
    Case TICK_BID
        crescendoTickType = FileTickTypes.Bid
    Case TICK_ASK
        crescendoTickType = FileTickTypes.Ask
    Case TICK_TRADE
        crescendoTickType = FileTickTypes.Last
    Case TICK_HIGH
        crescendoTickType = FileTickTypes.High
    Case TICK_LOW
        crescendoTickType = FileTickTypes.Low
    Case TICK_CLOSE
        crescendoTickType = FileTickTypes.PrevClose
    Case TICK_VOLUME
        crescendoTickType = FileTickTypes.Volume
    End Select
    writeTickfileLine CDbl(timestamp) & "," & _
                FormatDateTime(timestamp, vbGeneralDate) & "," & _
                crescendoTickType & "," & _
                price & "," & _
                size
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub writeTickfileLine( _
                ByVal text As String)
Const ProcName As String = "writeTickfileLine"
Dim failpoint As String
On Error GoTo Err

mTsOut.WriteLine text
mBytesWritten = mBytesWritten + Len(text)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub




