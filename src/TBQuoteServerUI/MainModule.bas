Attribute VB_Name = "MainModule"
Option Explicit

Public Declare Sub InitCommonControls Lib "comctl32" ()

Private mArguments As cCommandLineArgs
Private mForm As fDataCollectorUI

Private mClientID As Long
Private mShortNames() As String
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

Private mDataCollector As CTBDataCollector

Public Sub Main()
Dim posnValue As String
Dim i As Long

mLeftOffset = -1
mRightOffset = -1

Set mArguments = New cCommandLineArgs
mArguments.CommandLine = Command
mArguments.Separator = " "
mArguments.GetArgs

If mArguments.Switch("?") Then
    MsgBox vbCrLf & _
            "datacollector [/shortnames:symbolshortname[,symbolshortname]...]" & vbCrLf & _
            "              [/server:servername]" & vbCrLf & _
            "              [/clientid:clientidnumber]" & vbCrLf & _
            "              [/outformat:formatname]" & vbCrLf & _
            "              [/outpath:path]" & vbCrLf & _
            "              [/noWriteBars  |  /nwb]" & vbCrLf & _
            "              [/noWriteTicks  |  /nwt]" & vbCrLf & _
            "              [/posn:offsetfromleft,offsetfromtop]" & vbCrLf & _
            "              [/port:portnumber]" & vbCrLf & _
            "              [/noUI]", , "Usage"
    Exit Sub
End If

If mArguments.Switch("noui") Then
    mNoUI = True
End If

If mArguments.Switch("shortnames") Then
    Dim shortNamesStr As String
    shortNamesStr = mArguments.SwitchValue("shortnames")
    mShortNames = Split(shortNamesStr, ",")
    
    For i = 0 To UBound(mShortNames)
        mShortNames(i) = Trim$(Replace(mShortNames(i), """", ""))
    Next
End If


If mArguments.Switch("server") Then
    mServer = mArguments.SwitchValue("server")
End If

If mArguments.Switch("clientid") Then
    If IsNumeric(mArguments.SwitchValue("clientid")) Then
        mClientID = mArguments.SwitchValue("clientid")
    Else
        If mNoUI Then
            Exit Sub
        Else
            MsgBox "Error - clientid  " & mArguments.SwitchValue("clientid") & " is not numeric"
            Exit Sub
        End If
    End If
Else
    Randomize
    mClientID = CLng(&H7FFFFFFF * Rnd)
End If

If mArguments.Switch("nwb") Or _
    mArguments.Switch("nowritebars") _
Then
    mNoWriteBars = True
End If

If mArguments.Switch("nwt") Or _
    mArguments.Switch("nowriteticks") _
Then
    mNoWriteTicks = True
End If

If mArguments.Switch("outpath") Then
    mOutputTickfilePath = mArguments.SwitchValue("outpath")
End If

If mArguments.Switch("outformat") Then
    mOutputTickfileFormat = mArguments.SwitchValue("outformat")
End If

If mArguments.Switch("port") Then
    If IsNumeric(mArguments.SwitchValue("port")) Then
        mPort = mArguments.SwitchValue("port")
    Else
        If mNoUI Then
            Exit Sub
        Else
            MsgBox "Error - port  " & mArguments.SwitchValue("port") & " is not numeric"
            Exit Sub
        End If
    End If
Else
    mPort = 7496
End If

Set mDataCollector = New CTBDataCollector
mDataCollector.ShortNames = mShortNames
mDataCollector.Server = mServer
mDataCollector.Port = mPort
mDataCollector.ClientID = mClientID
mDataCollector.WriteBars = Not mNoWriteBars
mDataCollector.WriteTicks = Not mNoWriteTicks
mDataCollector.OutputPath = mOutputTickfilePath
mDataCollector.OutputFormat = mOutputTickfileFormat

If Not mNoUI Then
    Set mForm = New fDataCollectorUI
    
    mForm.dataCollector = mDataCollector
    
    If mArguments.Switch("posn") Then
        posnValue = mArguments.SwitchValue("posn")
        
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

mDataCollector.startCollection

End Sub



