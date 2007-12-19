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

Public gStop As Boolean

Private mCLParser As CommandLineParser
Private mForm As fDataCollectorUI

Private mConfigPath As String

Private mNoAutoStart As Boolean
Private mNoUI As Boolean
Private mLeftOffset As Long
Private mRightOffset As Long
Private mPosX As Single
Private mPosY As Single

Private mDataCollector As dataCollector

Private mStartTimeDescriptor As String
Private mEndTimeDescriptor As String
Private mExitTimeDescriptor As String

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

On Error GoTo Err

mLeftOffset = -1
mRightOffset = -1

Set mCLParser = CreateCommandLineParser(Command, " ")

If showHelp Then Exit Sub

InitialiseTWUtilities

mNoUI = getNoUi

mConfigPath = getConfig
mStartTimeDescriptor = getStartTimeDescriptor
mEndTimeDescriptor = getEndTimeDescriptor
mExitTimeDescriptor = getExitTimeDescriptor

mNoAutoStart = getNoAutostart

If mNoUI Then
    
    Set mDataCollector = CreateDataCollector(mConfigPath, _
                                            mStartTimeDescriptor, _
                                            mEndTimeDescriptor, _
                                            mExitTimeDescriptor)
    
    If mStartTimeDescriptor = "" Then
        mDataCollector.startCollection
    End If
    
    Do While Not gStop
        Wait 1000
    Loop
    
    TerminateTWUtilities
    
Else
    Set mDataCollector = CreateDataCollector(mConfigPath, _
                                            IIf(mNoAutoStart, "", mStartTimeDescriptor), _
                                            mEndTimeDescriptor, _
                                            mExitTimeDescriptor)
    
    Set mForm = createForm
End If


Exit Sub

Err:
MsgBox "Error " & Err.Number & ": " & Err.description & vbCrLf & _
        "At " & "TBQuoteServerUI" & "." & "MainModule" & "::" & "Main" & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        , _
        "Error"


End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createForm() As fDataCollectorUI
Dim posnValue As String

Set createForm = New fDataCollectorUI

If mCLParser.Switch("posn") Then
    posnValue = mCLParser.SwitchValue("posn")
    
    If InStr(1, posnValue, ",") = 0 Then
        MsgBox "Error - posn value must be 'n,m'"
        Set createForm = Nothing
        Exit Function
    End If
    
    If Not IsNumeric(Left$(posnValue, InStr(1, posnValue, ",") - 1)) Then
        MsgBox "Error - offset from left is not numeric"
        Set createForm = Nothing
        Exit Function
    End If
    
    mPosX = Left$(posnValue, InStr(1, posnValue, ",") - 1)
    
    If Not IsNumeric(Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))) Then
        MsgBox "Error - offset from top is not numeric"
        Set createForm = Nothing
        Exit Function
    End If
    
    mPosY = Right$(posnValue, Len(posnValue) - InStr(1, posnValue, ","))
Else
    mPosX = Int(Int(Screen.Width / createForm.Width) * Rnd)
    mPosY = Int(Int(Screen.Height / createForm.Height) * Rnd)
End If

createForm.initialise mDataCollector, _
                getNoAutostart, _
                mCLParser.Switch("showMonitor")

createForm.Left = mPosX * createForm.Width
createForm.Top = mPosY * createForm.Height

createForm.Visible = True
End Function

Private Function getConfig() As String

If mCLParser.Switch("Config") Then
    getConfig = mCLParser.SwitchValue("config")
Else
    getConfig = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                        "\TradeWright\" & _
                        App.EXEName & _
                        "\v" & _
                        App.Major & "." & App.Minor & _
                        "\settings.xml"
End If
End Function

Private Function getEndTimeDescriptor() As String
If mCLParser.Switch("endAt") Then
    getEndTimeDescriptor = mCLParser.SwitchValue("endAt")
End If
End Function

Private Function getExitTimeDescriptor() As String
If mCLParser.Switch("exitAt") Then
    getExitTimeDescriptor = mCLParser.SwitchValue("exitAt")
End If
End Function

Private Function getNoAutostart() As Boolean
If mCLParser.Switch("noAutoStart") Then
    getNoAutostart = True
End If
End Function

Private Function getNoUi() As Boolean
If mCLParser.Switch("noui") Then
    getNoUi = True
End If
End Function

Private Function getStartTimeDescriptor() As String
If mCLParser.Switch("startAt") Then
    getStartTimeDescriptor = mCLParser.SwitchValue("startAt")
End If
End Function

Private Function showHelp() As Boolean

If mCLParser.Switch("?") Or mCLParser.NumberOfSwitches = 0 Then
    MsgBox vbCrLf & _
            "datacollector26 [/config:filename] " & vbCrLf & _
            "              [/posn:offsetfromleft,offsetfromtop]" & vbCrLf & _
            "              [/noAutoStart" & vbCrLf & _
            "              [/noUI]" & vbCrLf & _
            "              [/showMonitor]" & vbCrLf & _
            "              [/exitAt:[day]hh:mm]" & vbCrLf & _
            "              [/startAt:[day]hh:mm]" & vbCrLf & _
            "              [/endAt:[day]hh:mm]", _
            , _
            "Usage"
    showHelp = True
End If
End Function

