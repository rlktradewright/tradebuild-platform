Attribute VB_Name = "MainModule"
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================


'@================================================================================
' Member variables
'@================================================================================

Private mCLParser As CommandLineParser
Private mForm As fDataCollectorUI

Private mFso As FileSystemObject
Private mSymbsTS As TextStream

Private mClientID As Long
Private mInstruments() As InstrumentSpecifier
Private mNumInstruments As Long

Private mShortNames() As String
Private mNumShortNames As Long

Private mServer As String
Private mOutputTickfilePath As String
Private mOutputTickfileFormat As String
Private mPort As Long
Private mNoWriteBars As Boolean
Private mNoWriteTicks As Boolean
Private mNoUI As Boolean
Private mLeftOffset As Long
Private mRightOffset As Long
Private mPosX As Single
Private mPosY As Single

Private mDataCollector As TBDataCollector

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Main()
Dim posnValue As String
Dim i As Long
Dim ar() As String
Dim rec As String

Dim failpoint As Long
On Error GoTo Err

mLeftOffset = -1
mRightOffset = -1

Set mCLParser = CreateCommandLineParser(Command, " ")

failpoint = 100 '----------------------------------------------------------

If mCLParser.Switch("?") Or mCLParser.NumberOfSwitches = 0 Then
    MsgBox vbCrLf & _
            "datacollector [/symbs:filename] " & vbCrLf & _
            "              [/shortnames:symbolshortname[,symbolshortname]...]" & vbCrLf & _
            "              [/server:servername]" & vbCrLf & _
            "              [/port:portnumber]" & vbCrLf & _
            "              [/clientid:clientidnumber]" & vbCrLf & _
            "              [/outformat:formatname]" & vbCrLf & _
            "              [/outpath:path]" & vbCrLf & _
            "              [/noWriteBars  |  /nwb]" & vbCrLf & _
            "              [/noWriteTicks  |  /nwt]" & vbCrLf & _
            "              [/posn:offsetfromleft,offsetfromtop]" & vbCrLf & _
            "              [/noUI]", , "Usage"
    Exit Sub
End If

failpoint = 200 '----------------------------------------------------------

If mCLParser.Switch("noui") Then
    mNoUI = True
End If

failpoint = 300 '----------------------------------------------------------

If mCLParser.Switch("symbs") Then
    
    ReDim mInstruments(15) As InstrumentSpecifier

    Set mFso = New FileSystemObject
    Set mSymbsTS = mFso.OpenTextFile(mCLParser.SwitchValue("symbs"))
    Do While Not mSymbsTS.AtEndOfStream
        If mNumInstruments = 16 Then
            MsgBox "Attempting to collect data for more than 16 instruments"
            Exit Sub
        End If
        rec = mSymbsTS.ReadLine
        If rec <> "" And Left$(rec, 2) <> "//" Then
            ar = Split(rec, ",")
            mInstruments(mNumInstruments).ShortName = Trim$(ar(0))
            mInstruments(mNumInstruments).symbol = Trim$(ar(1))
            mInstruments(mNumInstruments).secType = Trim$(ar(2))
            mInstruments(mNumInstruments).expiry = Trim$(ar(3))
            mInstruments(mNumInstruments).exchange = Trim$(ar(4))
            mInstruments(mNumInstruments).currencyCode = Trim$(ar(5))
            mInstruments(mNumInstruments).strikePrice = Trim$(ar(6))
            mInstruments(mNumInstruments).Right = Trim$(ar(7))
            mNumInstruments = mNumInstruments + 1
        End If
    Loop
    
    If mNumInstruments <> 0 Then
        ReDim Preserve mInstruments(mNumInstruments - 1) As InstrumentSpecifier
    End If
End If

failpoint = 400 '----------------------------------------------------------

If mCLParser.Switch("shortnames") Then
    Dim shortNamesStr As String
    shortNamesStr = mCLParser.SwitchValue("shortnames")
    mShortNames = Split(shortNamesStr, ",")
    
    For i = 0 To UBound(mShortNames)
        mShortNames(i) = Trim$(Replace(mShortNames(i), """", ""))
    Next
    
    mNumShortNames = UBound(mShortNames) + 1
End If

If mNumInstruments + mNumShortNames > 16 Then
    MsgBox "Attempting to collect data for more than 16 instruments"
    Exit Sub
End If
    

failpoint = 500 '----------------------------------------------------------

If mCLParser.Switch("server") Then
    mServer = mCLParser.SwitchValue("server")
End If

If mCLParser.Switch("clientid") Then
    If IsNumeric(mCLParser.SwitchValue("clientid")) Then
        mClientID = mCLParser.SwitchValue("clientid")
    Else
        If mNoUI Then
            Exit Sub
        Else
            MsgBox "Error - clientid  " & mCLParser.SwitchValue("clientid") & " is not numeric"
            Exit Sub
        End If
    End If
Else
    Randomize
    mClientID = CLng(&H7FFFFFFF * Rnd)
End If

failpoint = 600 '----------------------------------------------------------

If mCLParser.Switch("nwb") Or _
    mCLParser.Switch("nowritebars") _
Then
    mNoWriteBars = True
End If

failpoint = 700 '----------------------------------------------------------

If mCLParser.Switch("nwt") Or _
    mCLParser.Switch("nowriteticks") _
Then
    mNoWriteTicks = True
End If

failpoint = 800 '----------------------------------------------------------

If mCLParser.Switch("outpath") Then
    mOutputTickfilePath = mCLParser.SwitchValue("outpath")
End If

failpoint = 900 '----------------------------------------------------------

If mCLParser.Switch("outformat") Then
    mOutputTickfileFormat = mCLParser.SwitchValue("outformat")
End If

failpoint = 1000 '---------------------------------------------------------

If mCLParser.Switch("port") Then
    If IsNumeric(mCLParser.SwitchValue("port")) Then
        mPort = mCLParser.SwitchValue("port")
    Else
        If mNoUI Then
            Exit Sub
        Else
            MsgBox "Error - port  " & mCLParser.SwitchValue("port") & " is not numeric"
            Exit Sub
        End If
    End If
Else
    mPort = 7496
End If

failpoint = 1100 '---------------------------------------------------------

Set mDataCollector = New TBDataCollector
mDataCollector.ShortNames = mShortNames
mDataCollector.Instruments = mInstruments
mDataCollector.Server = mServer
mDataCollector.Port = mPort
mDataCollector.ClientID = mClientID
mDataCollector.WriteBars = Not mNoWriteBars
mDataCollector.WriteTicks = Not mNoWriteTicks
If mOutputTickfilePath = "" Then
    mDataCollector.OutputPath = App.Path & "\Tickfiles"
Else
    mDataCollector.OutputPath = mOutputTickfilePath
End If
If mOutputTickfileFormat = "" Then
    mDataCollector.OutputFormat = "TradeBuild V4"
Else
    mDataCollector.OutputFormat = mOutputTickfileFormat
End If

failpoint = 1200 '---------------------------------------------------------

If Not mNoUI Then
    Set mForm = New fDataCollectorUI
    
    mForm.dataCollector = mDataCollector
    
    If mCLParser.Switch("posn") Then
        posnValue = mCLParser.SwitchValue("posn")
        
        If InStr(1, posnValue, ",") = 0 Then
            MsgBox "Error - posn value must be 'n,m'"
            Exit Sub
        End If
        
        If Not IsNumeric(Left$(posnValue, InStr(1, posnValue, ",") - 1)) Then
            MsgBox "Error - offset from left is not numeric"
            Exit Sub
        End If
        
        mPosX = Left$(posnValue, InStr(1, posnValue, ",") - 1)
        
        If Not IsNumeric(Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))) Then
            MsgBox "Error - offset from top is not numeric"
            Exit Sub
        End If
        
        mPosY = Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))
    Else
        Randomize
        mPosX = Int(Int(Screen.Width / mForm.Width) * Rnd)
        Randomize
        mPosY = Int(Int(Screen.Height / mForm.Height) * Rnd)
    End If
    
    mForm.Left = mPosX * mForm.Width
    mForm.Top = mPosY * mForm.Height
    mForm.Visible = True
    
End If

failpoint = 1200 '---------------------------------------------------------

mDataCollector.startCollection

Exit Sub

Err:
MsgBox "Error " & Err.Number & ": " & Err.description & vbCrLf & _
        "At " & "TBQuoteServerUI" & "." & "MainModule" & "::" & "Main" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        , _
        "Error"


End Sub

'@================================================================================
' Helper Functions
'@================================================================================


