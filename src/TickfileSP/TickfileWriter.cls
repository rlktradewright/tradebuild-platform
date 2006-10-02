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

Implements ITickfileWriter

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickfileServiceProvider As tickfileServiceProvider

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mServiceConsumer As ITickfileOutputServiceConsumer

Private mFileSys As Scripting.FileSystemObject
Private mTsOut As Scripting.TextStream

Private mTickerID As Long
Private mContract As IContract

Private mPath As String

Private mTickFileVersion As TickFileVersions
Private mTickfileFormat As TickfileFormats

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mPath = App.path
If Right$(mPath, 1) <> "\" Then mPath = mPath & "\"
Set mFileSys = CreateObject("scripting.filesystemobject")
End Sub

Private Sub Class_Terminate()
Debug.Print "TickfileSP.TickfileWriter terminated"
End Sub

'================================================================================
' ITickfileWriter Interface Members
'================================================================================

Private Sub ITickfileWriter_CloseOutputFile()
On Error GoTo Err
finish
mTickfileServiceProvider.dataWriterFinished Me
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_CloseOutputFile", _
                Err.Description
End Sub

Private Property Get ITickfileWriter_Contract() As IContract
Set ITickfileWriter_Contract = mContract
End Property

Private Property Let ITickfileWriter_Contract(ByVal RHS As IContract)
If Not RHS Is Nothing Then
    Set mContract = RHS.Clone
End If
mServiceConsumer.Ready = True
End Property

Private Sub ITickfileWriter_WriteAsk(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
On Error GoTo Err
writeTickData timestamp, TICK_ASK, price, size
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_WriteAsk", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeBid(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
On Error GoTo Err
writeTickData timestamp, TICK_BID, price, size
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeBid", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeClose(ByVal timestamp As Date, ByVal price As Double)
On Error GoTo Err
writeTickData timestamp, TICK_CLOSE, price
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeClose", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeHigh(ByVal timestamp As Date, ByVal price As Double)
On Error GoTo Err
writeTickData timestamp, TICK_HIGH, price
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeHigh", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeLow(ByVal timestamp As Date, ByVal price As Double)
writeTickData timestamp, TICK_LOW, price
End Sub

Private Sub ITickfileWriter_writeMarketDepthData(ByVal timestamp As Date, ByVal Position As Long, ByVal marketMaker As String, ByVal operation As Long, ByVal side As Long, ByVal price As Double, ByVal size As Long)
On Error GoTo Err
Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    If mTsOut Is Nothing Then createTradeBuildTickFile timestamp
    
    mTsOut.WriteLine CDbl(timestamp) & "," & _
                mCommonServiceConsumer.FormatTimestamp(timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                TICK_MARKET_DEPTH & "," & _
                Position & "," & _
                marketMaker & "," & _
                operation & "," & _
                side & "," & _
                format(price, "Fixed") & "," & _
                size
Case TickfileFormats.TickfileCrescendo
    ' Crescendo can't handle market depth in tickfiles
End Select
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthData", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeMarketDepthReset(ByVal timestamp As Date)
On Error GoTo Err
Select Case mTickfileFormat
Case TickfileFormats.TickfileTradeBuild
    If mTsOut Is Nothing Then createTradeBuildTickFile timestamp
    
    mTsOut.WriteLine CDbl(timestamp) & "," & _
                mCommonServiceConsumer.FormatTimestamp(timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                TICK_MARKET_DEPTH_RESET

Case TickfileFormats.TickfileCrescendo
    ' Crescendo can't handle market depth in tickfiles
End Select
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeMarketDepthReset", _
                Err.Description
End Sub

Private Sub ITickfileWriter_WriteOpenInterest(ByVal timestamp As Date, ByVal size As Long)
On Error GoTo Err
writeTickData timestamp, TICK_OPEN_INTEREST, , size
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_WriteOpenInterest", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeTrade(ByVal timestamp As Date, ByVal price As Double, ByVal size As Long)
On Error GoTo Err
writeTickData timestamp, TICK_TRADE, price, size
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeTrade", _
                Err.Description
End Sub

Private Sub ITickfileWriter_writeVolume(ByVal timestamp As Date, ByVal size As Long)
On Error GoTo Err
writeTickData timestamp, TICK_VOLUME, , size
Exit Sub
Err:
handleFatalError Err.Number, _
                "TickfileSP" & "." & "TickfileWriter" & "::" & "ITickfileWriter_writeVolume", _
                Err.Description
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Let commonServiceConsumer(ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Friend Property Let path(ByVal value As String)
If value = "" Then Exit Property
mPath = value
If Right$(mPath, 1) <> "\" Then mPath = mPath & "\"
End Property

Friend Property Let tickerid(ByVal value As Long)
mTickerID = value
End Property

Friend Property Get TickfileFormat() As TickfileFormats
TickfileFormat = mTickfileFormat
End Property

Friend Property Let TickfileFormat(ByVal value As TickfileFormats)
mTickfileFormat = value
End Property

Friend Property Let TickfileOutputServiceConsumer(ByVal RHS As ITickfileOutputServiceConsumer)
Set mServiceConsumer = RHS
End Property

Friend Property Let tickfileServiceProvider( _
                ByVal value As tickfileServiceProvider)
Set mTickfileServiceProvider = value
End Property

Friend Property Get TickFileVersion() As TickFileVersions
TickFileVersion = mTickFileVersion
End Property

Friend Property Let TickFileVersion(ByVal value As TickFileVersions)
mTickFileVersion = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub finish()
If Not mTsOut Is Nothing Then
    mTsOut.Close
    Set mTsOut = Nothing
End If
Set mServiceConsumer = Nothing
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub createCrescendoTickfile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String
Const TICKFILE_CURR_VERSION As Integer = 2

On Error GoTo Err

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"

With mContract.Specifier
    tfPath = tfPath & UCase$(.Symbol) & .Expiry & "\"
    tfName = UCase$(.Symbol) & .Expiry & "-" _
            & format(timestamp, "yyyymmdd-hhmmss") & ".csv"
    Set mTsOut = mFileSys.CreateTextFile(tfPath & tfName)
    
    mServiceConsumer.OutputTickfileCreated tfPath & tfName
    
    mTsOut.WriteLine "tickfile," & _
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
Else
    Err.Raise Err.Number, , Err.Description
End If
End Sub

Private Sub createTradeBuildTickFile(ByVal timestamp As Date)
Dim tfPath As String
Dim tfName As String

On Error GoTo Err

tfPath = IIf(mPath <> "", mPath, App.path)
If Right$(tfPath, 1) <> "\" Then tfPath = tfPath & "\"
tfPath = tfPath & UCase$(mContract.Specifier.LocalSymbol) & "\"
tfName = UCase$(mContract.Specifier.LocalSymbol) & "-" & _
        format(timestamp, "yyyymmdd-hhmmss") & ".tck"
Set mTsOut = mFileSys.CreateTextFile(tfPath & tfName)
mServiceConsumer.OutputTickfileCreated tfPath & tfName
mTsOut.WriteLine TICKFILE_DECLARER & "," & _
            TICKFILE_CURR_VERSION & "," & _
            UCase$(mContract.Specifier.Exchange) & "," & _
            UCase$(mContract.Specifier.Symbol) & "," & _
            mContract.Specifier.Expiry & "," & _
            CDbl(timestamp) & "," & _
            FormatDateTime(timestamp, vbGeneralDate)
mTsOut.WriteLine CONTRACT_DETAILS_MARKER & mContract.ToXML

Exit Sub

Err:
If Err.Number = 76 Then
    mFileSys.CreateFolder tfPath
    Resume
Else
    Err.Raise Err.Number, , Err.Description
End If
End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, _
                            Source, _
                            Description, _
                            mTickfileServiceProvider.handle

finish
mTickfileServiceProvider.dataWriterFinished Me
End Sub

Private Sub writeTickData( _
                        ByVal timestamp As Date, _
                        ByVal tickType As String, _
                        Optional ByVal price As Double = -1#, _
                        Optional ByVal size As Long = -1&)
Dim crescendoTickType As TickTypes
Static priceFormatString As String


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
    mTsOut.WriteLine CDbl(timestamp) & "," & _
                mCommonServiceConsumer.FormatTimestamp(timestamp, TimestampFormats.TimestampTimeOnly) & "," & _
                tickType & "," & _
                IIf(price <> -1#, format(price, priceFormatString), "") & _
                IIf(size = -1&, "", IIf(price = -1#, size, "," & size))
Case TickfileFormats.TickfileCrescendo
    If mTsOut Is Nothing Then createCrescendoTickfile timestamp
    Select Case tickType
    Case TICK_BID
        crescendoTickType = TickTypes.Bid
    Case TICK_ASK
        crescendoTickType = TickTypes.Ask
    Case TICK_TRADE
        crescendoTickType = TickTypes.Last
    Case TICK_HIGH
        crescendoTickType = TickTypes.High
    Case TICK_LOW
        crescendoTickType = TickTypes.Low
    Case TICK_CLOSE
        crescendoTickType = TickTypes.PrevClose
    Case TICK_VOLUME
        crescendoTickType = TickTypes.Volume
    End Select
    mTsOut.WriteLine CDbl(timestamp) & "," & _
                FormatDateTime(timestamp, vbGeneralDate) & "," & _
                crescendoTickType & "," & _
                price & "," & _
                size
End Select
End Sub






