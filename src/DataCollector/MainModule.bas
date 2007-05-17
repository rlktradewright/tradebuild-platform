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

Private mConfigPath As String

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

InitialiseTWUtilities
InitialiseTimerUtils

mLeftOffset = -1
mRightOffset = -1

Set mCLParser = CreateCommandLineParser(Command, " ")

failpoint = 100 '----------------------------------------------------------

If mCLParser.Switch("?") Or mCLParser.NumberOfSwitches = 0 Then
    MsgBox vbCrLf & _
            "datacollector26 [/config:filename] " & vbCrLf & _
            "              [/posn:offsetfromleft,offsetfromtop]" & vbCrLf & _
            "              [/noUI]", , "Usage"
    Exit Sub
End If

failpoint = 200 '----------------------------------------------------------

If mCLParser.Switch("noui") Then
    mNoUI = True
End If

failpoint = 300 '----------------------------------------------------------

If mCLParser.Switch("Config") Then
    mConfigPath = mCLParser.SwitchValue("config")
Else
    mConfigPath = GetSpecialFolderPath(FolderIdLOCAL_APPDATA) & _
                        "\TradeWright\" & _
                        App.EXEName & _
                        "\v" & _
                        App.Major & "." & App.Minor & _
                        "\settings.xml"
End If

failpoint = 400 '----------------------------------------------------------

Set mDataCollector = CreateDataCollector(mConfigPath)

failpoint = 1200 '---------------------------------------------------------

If mNoUI Then
    failpoint = 1400 '---------------------------------------------------------
    
    mDataCollector.startCollection
    
    Do
        Wait 1000
    Loop
    
    TerminateTWUtilities
    TerminateTimerUtils
Else
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
    
    failpoint = 1300 '---------------------------------------------------------
    
    mDataCollector.startCollection
End If


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


