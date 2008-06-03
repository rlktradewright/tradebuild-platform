Attribute VB_Name = "GMain"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeSkilDemo26"
Private Const ModuleName                    As String = "GMain"

'@================================================================================
' Member variables
'@================================================================================

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

Public Sub main()
Dim lTradeSkilDemo As fTradeSkilDemo

InitialiseTWUtilities
DefaultLogLevel = TWUtilities30.LogLevels.LogLevelNormal

If showCommandLineOptions() Then Exit Sub

If Not getLog() Then Exit Sub

Set lTradeSkilDemo = New fTradeSkilDemo
If lTradeSkilDemo.configure Then
    lTradeSkilDemo.Show vbModeless
Else
    Unload lTradeSkilDemo
End If

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getLog() As Boolean
Dim listener As LogListener

On Error GoTo Err

If gCommandLineParser.Switch(SwitchLogLevel) Then DefaultLogLevel = LogLevelFromString(gCommandLineParser.SwitchValue(SwitchLogLevel))

Set listener = CreateFileLogListener(gLogFileName, _
                                        CreateBasicLogFormatter, _
                                        True, _
                                        False)
' ensure log entries of all infotypes get written to the log file
gLogger.addLogListener listener

getLog = True
Exit Function

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to  '" & gLogFileName & "': the program will close", vbCritical, "Attention"
    getLog = False
Else
    Err.Raise Err.Number    ' unknown error so re-raise it
End If
End Function

Private Function showCommandLineOptions() As Boolean

If gCommandLineParser.Switch("?") Then
    MsgBox vbCrLf & _
            "tradeskildemo26 [configfile] " & vbCrLf & _
            "                [/config:configtoload] " & vbCrLf & _
            "                [/log:filename] " & vbCrLf & _
            "                [/loglevel:levelName]" & vbCrLf & _
            vbCrLf & _
            "  where" & vbCrLf & _
            vbCrLf & _
            "    levelname is one of:" & vbCrLf & _
            "       None    or 0" & vbCrLf & _
            "       Severe  or S" & vbCrLf & _
            "       Warning or W" & vbCrLf & _
            "       Info    or I" & vbCrLf & _
            "       Normal  or N" & vbCrLf & _
            "       Detail  or D" & vbCrLf & _
            "       Medium  or M" & vbCrLf & _
            "       High    or H" & vbCrLf & _
            "       All     or A", _
            , _
            "Usage"
    showCommandLineOptions = True
Else
    showCommandLineOptions = False
End If
End Function


