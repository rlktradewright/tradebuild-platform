Attribute VB_Name = "MainModule"
Option Explicit

Private Const AppName As String = "ChartDemo"

Public Sub Main()
If getLog Then ChartForm.Show
End Sub

Private Function getLog() As Boolean
Dim listener As LogListener
Dim logFilename As String

On Error GoTo Err

DefaultLogLevel = TWUtilities30.LogLevels.LogLevelHighDetail

logFilename = GetSpecialFolderPath(FolderIdLocalAppdata) & _
                                    "\TradeWright\" & _
                                    AppName & _
                                    "\v" & _
                                    App.Major & "." & App.Minor & _
                                    "\log.txt"
Set listener = CreateFileLogListener(logFilename, _
                                        CreateBasicLogFormatter, _
                                        True, _
                                        False)
' ensure log entries of all infotypes get written to the log file
GetLogger("").addLogListener listener

getLog = True
Exit Function

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    MsgBox "You don't have write access to  '" & logFilename & "': the program will close", vbCritical, "Attention"
    getLog = False
Else
    Err.Raise Err.Number    ' unknown error so re-raise it
End If
End Function



