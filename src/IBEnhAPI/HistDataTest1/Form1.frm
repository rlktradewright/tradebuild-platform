VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#377.0#0"; "TradingUI27.ocx"
Begin VB.Form Form1 
   Caption         =   "Historical Data Tester"
   ClientHeight    =   10380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   LinkTopic       =   "Form1"
   ScaleHeight     =   10380
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox DisableRequestPacingCheck 
      Caption         =   "Disable request pacing"
      Height          =   495
      Left            =   8880
      TabIndex        =   9
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CheckBox ShowBarOnReceiptCheck 
      Caption         =   "Show each bar on receipt"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   3360
      Width           =   1575
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6191
      ForeColor       =   -2147483640
      ModeAdvanced    =   -1  'True
   End
   Begin VB.CommandButton FetchYearsDataButton 
      Caption         =   "Fetch year's data"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   5
      Top             =   2040
      Width           =   1575
   End
   Begin VB.CheckBox ShowBarsAtEndCheck 
      Caption         =   "Show bars when load complete"
      Height          =   495
      Left            =   8880
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton FetchConstRangeButton 
      Caption         =   "Fetch const range"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   4
      Top             =   1560
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
      TabIndex        =   8
      Top             =   4560
      Width           =   10335
   End
   Begin VB.CommandButton BasicTestButton 
      Caption         =   "Basic fetch test"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8880
      TabIndex        =   1
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

Private mClient                                     As Client
Attribute mClient.VB_VarHelpID = -1
Private mContractStore                              As IContractStore
Private mHistDataStore                              As IHistoricalDataStore

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
Set mUnhandledErrorHandler = UnhandledErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "HistDataTest1"
DefaultLogLevel = LogLevelHighDetail
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

mClientId = 74889561
Set mClient = GetClient("Essy", 7497, mClientId, , , ApiMessageLoggingOptionAlways, ApiMessageLoggingOptionNone, False, , Me)
Set mContractStore = mClient.GetContractStore
Set mHistDataStore = mClient.GetHistoricalDataStore
If DisableRequestPacingCheck.Value = vbChecked Then mClient.DisableHistoricalDataRequestPacing

End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Select Case pState
Case ApiConnNotConnected
    disableControls
    LogMessage "Disconnected from TWS: " & pMessage
Case ApiConnConnecting
    LogMessage "Connecting to TWS: " & pMessage
Case ApiConnConnected
    enableControls
    If ContractSpecBuilder1.IsReady Then FetchYearsDataButton.Enabled = True
    LogMessage "Connected to TWS: " & pMessage
Case ApiConnFailed
    disableControls
    LogMessage "Failed to connect to TWS: " & pMessage
End Select
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

Dim lBarSpec As BarDataSpecifier
Dim lContractSpec As IContractSpecifier

Set lContractSpec = CreateContractSpecifier("ESM2", "ES", "GLOBEX", SecTypeFuture, "USD", "202206")

Set lBarSpec = CreateBarDataSpecifier( _
                GetTimePeriod(1, TimePeriodMinute), _
                CDate("2022/05/25 10:00"), _
                CDate("2022/05/25 13:00"), _
                10000, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("17:30"))

FetchBars FetchContract(lContractSpec, mContractStore, pCookie:=lContractSpec), lBarSpec, lContractSpec

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSpecBuilder1_NotReady()
FetchYearsDataButton.Enabled = False
End Sub

Private Sub ContractSpecBuilder1_Ready()
If mClient.TwsApiConnectionState = ApiConnConnected Then FetchYearsDataButton.Enabled = True
End Sub

Private Sub FetchConstRangeButton_Click()
Const ProcName As String = "FetchConstRangeButton_Click"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier("ESM2", "ES", "GLOBEX", SecTypeFuture, "USD", "202206")

Dim lBarSpec As BarDataSpecifier
Set lBarSpec = CreateBarDataSpecifier( _
                GetTimePeriod(10, TimePeriodTickMovement), _
                CDate("2022/05/25 10:00"), _
                CDate("2022/05/25 13:00"), _
                500, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("17:30"))

FetchBars FetchContract(lContractSpec, mContractStore, pCookie:=lContractSpec), lBarSpec, lContractSpec

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FetchYearsDataButton_Click()
Const ProcName As String = "FetchYearsDataButton_Click"
On Error GoTo Err

Dim lBarSpec As BarDataSpecifier
Dim lContractSpec As IContractSpecifier

Set lContractSpec = ContractSpecBuilder1.ContractSpecifier

Set lBarSpec = CreateBarDataSpecifier( _
                GetTimePeriod(1, TimePeriodMinute), _
                CDate(Int(Now - 365#)), _
                Now, _
                365& * 1440&, _
                BarTypeTrade, _
                , _
                , _
                True)

FetchBars FetchContract(lContractSpec, mContractStore, pCookie:=lContractSpec), lBarSpec, lContractSpec

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton1_Click()
Const ProcName As String = "ManyFetchButton1_Click"
On Error GoTo Err

Dim lSymbols As New EnumerableCollection
lSymbols.Add "AAL"
lSymbols.Add "ABF"
lSymbols.Add "ADM"
lSymbols.Add "AGK"
lSymbols.Add "AMEC"
lSymbols.Add "ANTO"
lSymbols.Add "ARM"
lSymbols.Add "AU."
lSymbols.Add "AV."
lSymbols.Add "AZN"
' 10
lSymbols.Add "BA."
lSymbols.Add "BARC"
lSymbols.Add "BATS"
lSymbols.Add "BG."
lSymbols.Add "BLND"
lSymbols.Add "BLT"
lSymbols.Add "BP."
lSymbols.Add "BRBY"
lSymbols.Add "BSY"
lSymbols.Add "BT.A"
'20
lSymbols.Add "CCL"
lSymbols.Add "CNA"
lSymbols.Add "CNE"
lSymbols.Add "CPG"
lSymbols.Add "CPI"
lSymbols.Add "CSCG"
lSymbols.Add "DGE"
lSymbols.Add "EMG"
lSymbols.Add "ENRC"
lSymbols.Add "ESSR"
'30
lSymbols.Add "EXPN"
lSymbols.Add "FRES"
lSymbols.Add "GFS"
lSymbols.Add "GKN"
lSymbols.Add "GSK"
lSymbols.Add "HL."
lSymbols.Add "HMSO"
lSymbols.Add "HSBA"
lSymbols.Add "IAG"
lSymbols.Add "IAP"
'40
lSymbols.Add "IHG"
lSymbols.Add "III"
lSymbols.Add "IMI"
lSymbols.Add "IMT"
lSymbols.Add "INVP"
lSymbols.Add "IPR"
lSymbols.Add "ISAT"
lSymbols.Add "ISYS"
lSymbols.Add "ITRK"
lSymbols.Add "ITV"
'50

Dim t As New FetchFTSEBarsTask
t.Initialise Me, lSymbols

StartTask t, PriorityNormal

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton2_Click()
Const ProcName As String = "ManyFetchButton2_Click"
On Error GoTo Err

Dim lSymbols As New EnumerableCollection
lSymbols.Add "JMAT"
lSymbols.Add "KAZ"
lSymbols.Add "KGF"
lSymbols.Add "LAND"
lSymbols.Add "LGEN"
lSymbols.Add "LLOY"
lSymbols.Add "LMI"
lSymbols.Add "MKS"
lSymbols.Add "MRW"
lSymbols.Add "NG."
'60
lSymbols.Add "NXT"
lSymbols.Add "OML"
lSymbols.Add "PFC"
lSymbols.Add "PRU"
lSymbols.Add "PSON"
lSymbols.Add "RB."
lSymbols.Add "RBS"
lSymbols.Add "RDSA"
lSymbols.Add "RDSB"
lSymbols.Add "REL"
'70
lSymbols.Add "REX"
lSymbols.Add "RIO"
lSymbols.Add "RR."
lSymbols.Add "RRS"
lSymbols.Add "RSA"
lSymbols.Add "RSL"
lSymbols.Add "SAB"
lSymbols.Add "SBRY"
lSymbols.Add "SDR"
lSymbols.Add "SDRC"
'80
lSymbols.Add "SGE"
lSymbols.Add "SHP"
lSymbols.Add "SL."
lSymbols.Add "SMIN"
lSymbols.Add "SN."
lSymbols.Add "SRP"
lSymbols.Add "SSE"
lSymbols.Add "STAN"
lSymbols.Add "SVT"
lSymbols.Add "TLW"
'90
lSymbols.Add "TSCO"
lSymbols.Add "TT."
lSymbols.Add "ULVR"
lSymbols.Add "UU."
lSymbols.Add "VED"
lSymbols.Add "VOD"
lSymbols.Add "WEIR"
lSymbols.Add "WG."
lSymbols.Add "WOS"
lSymbols.Add "WPP"
'100
lSymbols.Add "WTB"
lSymbols.Add "XTA"

Dim t As New FetchFTSEBarsTask
t.Initialise Me, lSymbols

StartTask t, PriorityNormal

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

If Not mClient Is Nothing Then mClient.Finish

handleFatalError

End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function FetchBarsForFTSEStock(ByVal pSymbol As String) As IFuture
Const ProcName As String = "FetchBarsForFTSEStock"
On Error GoTo Err

Dim lBarSpec As BarDataSpecifier
Dim lContractSpec As IContractSpecifier

Set lContractSpec = CreateContractSpecifier(pSymbol, , "LSE", SecTypeStock, "GBP")

Set lBarSpec = CreateBarDataSpecifier( _
                GetTimePeriod(20, TimePeriodSecond), _
                Now - 7#, _
                Now, _
                2000, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("16:30"))

Set FetchBarsForFTSEStock = FetchBars(FetchContract(lContractSpec, mContractStore, pCookie:=lContractSpec), lBarSpec, lContractSpec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub disableControls()
BasicTestButton.Enabled = False
FetchConstRangeButton.Enabled = False
FetchYearsDataButton.Enabled = False
ManyFetchButton1.Enabled = False
ManyFetchButton2.Enabled = False
End Sub

Private Sub enableControls()
BasicTestButton.Enabled = True
FetchConstRangeButton.Enabled = True
ManyFetchButton1.Enabled = True
ManyFetchButton2.Enabled = True
End Sub

Private Function FetchBars(ByVal pContractFuture As IFuture, ByVal pBarSpec As BarDataSpecifier, pCookie As Variant) As IFuture
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lFetcher As New BarFetcher
Set FetchBars = lFetcher.Fetch( _
                    pBarSpec, _
                    mHistDataStore, _
                    pContractFuture, _
                    (ShowBarOnReceiptCheck.Value = vbChecked), _
                    (ShowBarsAtEndCheck.Value = vbChecked), _
                    pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

'If mIsInDev Then
'    End
'Else
'    EndProcess
'End If

End Sub

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function


