VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Contract Data Tester"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox ShowContractsCheck 
      Caption         =   "Show contracts"
      Height          =   255
      Left            =   8880
      TabIndex        =   4
      Top             =   2160
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton ManyFetchButton2 
      Caption         =   "Many fetches test 2"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.CommandButton ManyFetchButton1 
      Caption         =   "Many fetches test 1"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox LogText 
      Height          =   5775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   4560
      Width           =   10335
   End
   Begin VB.CommandButton BasicTestButton 
      Caption         =   "Basic fetch test"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ITwsConnectionStateListener
Implements ILogListener

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

Private Const ModuleName                            As String = "Form1"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mClientId                                   As Long

Private mResultCount                                As Long

Private mClient                                     As Client
Attribute mClient.VB_VarHelpID = -1

Private mContractStore                              As IContractStore

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
Set mUnhandledErrorHandler = UnhandledErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "ContractDataTest1"
DefaultLogLevel = LogLevelHighDetail
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Set mFutureWaiter = New FutureWaiter

mClientId = 25514278
Set mClient = GetClient("Essy", 7497, mClientId, , , ApiMessageLoggingOptionAlways, ApiMessageLoggingOptionNone, False, , Me)

Set mContractStore = mClient.GetContractStore

End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As IBENHAPI27.ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    disableControls
    LogMessage "Disconnected from TWS: " & pMessage
Case ApiConnConnecting
    LogMessage "Connecting to TWS: " & pMessage
Case ApiConnConnected
    enableControls
    LogMessage "Connected to TWS: " & pMessage
Case ApiConnFailed
    disableControls
    LogMessage "Failed to connect to TWS: " & pMessage
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)

End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)

End Sub

'@================================================================================
' ILogListener Interface Members
'@================================================================================

Private Sub ILogListener_Finish()

End Sub

Private Sub ILogListener_Notify(ByVal Logrec As LogRecord)
Const ProcName As String = "ILogListener_Notify"
On Error GoTo Err

If Len(LogText.Text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = formatLogRecord(Logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub BasicTestButton_Click()
Const ProcName As String = "BasicTestButton_Click"
On Error GoTo Err

Dim lSpec As IContractSpecifier
Set lSpec = CreateContractSpecifier(, "BA.", , "LSE", SecTypeStock, "GBP")
mFutureWaiter.Add mContractStore.FetchContracts(lSpec, , lSpec)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton1_Click()
Const ProcName As String = "ManyFetchButton1_Click"
On Error GoTo Err

Dim lSpec As IContractSpecifier
Set lSpec = CreateContractSpecifier(, "ES", , , , "USD")
mFutureWaiter.Add mContractStore.FetchContracts(lSpec, , lSpec)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton2_Click()
Const ProcName As String = "ManyFetchButton2_Click"
On Error GoTo Err

Dim lSpec As IContractSpecifier
Set lSpec = CreateContractSpecifier(, "ES", , , , "USD")
mFutureWaiter.Add mContractStore.FetchContracts(lSpec, , lSpec)

Set lSpec = CreateContractSpecifier(, "Z")
mFutureWaiter.Add mContractStore.FetchContracts(lSpec, , lSpec)

Set lSpec = CreateContractSpecifier(, "MSFT", , , , "USD")
mFutureWaiter.Add mContractStore.FetchContracts(lSpec, , lSpec)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

mResultCount = mResultCount + 1

Dim lSpec As IContractSpecifier
Set lSpec = ev.Future.Cookie

If ev.Future.IsFaulted Then
    LogMessage "(" & mResultCount & ") Error " & ev.Future.ErrorNumber & " for " & lSpec.ToString & vbCrLf & _
                        ev.Future.ErrorMessage & vbCrLf & _
                        ev.Future.ErrorSource
ElseIf ev.Future.IsCancelled Then
    LogMessage "(" & mResultCount & ") Contract details fetch cancelled for " & lSpec.ToString
Else
    Dim lContracts As IContracts
    Set lContracts = ev.Future.Value
    LogMessage "(" & mResultCount & ") Contract details fetch completed: " & lContracts.Count & " contracts retrieved for " & lSpec.ToString
    If ShowContractsCheck.Value = vbChecked Then
        Dim i As Long
        For i = 1 To lContracts.Count
            LogMessage "(" & mResultCount & "." & i & "): " & lContracts.ItemAtIndex(i).ToString
        Next
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

handleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
UnhandledErrorHandler.Handled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub disableControls()
BasicTestButton.Enabled = False
ManyFetchButton1.Enabled = False
ManyFetchButton2.Enabled = False
End Sub

Private Sub enableControls()
BasicTestButton.Enabled = True
ManyFetchButton1.Enabled = True
ManyFetchButton2.Enabled = True
End Sub

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
Static formatter As ILogFormatter

On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function




