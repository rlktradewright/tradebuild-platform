Attribute VB_Name = "MainModule"
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "ChartDemo26"
Private Const ModuleName                            As String = "MainModule"

Private Const AppName As String = "ChartDemo"

'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev As Boolean

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

Public Property Get gAppTitle() As String
gAppTitle = AppName & _
                " v" & _
                App.Major & "." & App.Minor & "." & App.Revision
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gHandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly. Note that normally one would use the
' End statement to terminate a VB6 program abruptly. However the TWUtilities component interferes
' with the End statement's processing and prevents proper shutdown, so we use the
' TWUtilities component's EndProcess method instead. (However if we are running in the
' development environment, then we call End because the EndProcess method kills the
' entire development environment as well which can have undesirable side effects if other
' components are also loaded.)

If mIsInDev Then
    ' this tells TWUtilities that we've now handled this unhandled error. Not actually
    ' needed here because the End statement will prevent return to TWUtilities
    UnhandledErrorHandler.Handled = True
    End
Else
    EndProcess
End If

End Sub

Public Sub Main()
Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
ApplicationGroupName = "TradeWright"
ApplicationName = gAppTitle
SetupDefaultLogging Command
ChartForm.Show
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function







