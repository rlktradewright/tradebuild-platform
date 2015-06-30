VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#292.0#0"; "TradingUI27.ocx"
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
   Begin VB.CheckBox ShowBarOnReceipt 
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
      _ExtentY        =   5556
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
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

mClientId = 74889561
Set mClient = GetClient("Sven", 7497, mClientId, , , , Me)
Set mContractStore = mClient.GetContractStore
Set mHistDataStore = mClient.GetHistoricalDataStore

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

Dim lBarSpecFuture As IFuture
Dim lContractSpec As IContractSpecifier

Set lContractSpec = CreateContractSpecifier("ZZ3", "Z", "LIFFE", SecTypeFuture, "GBP", "201312")

Set lBarSpecFuture = CreateBarDataSpecifierFuture(FetchContract(lContractSpec, mContractStore), _
                GetTimePeriod(1, TimePeriodMinute), _
                CDate("2013/11/01 08:00"), _
                CDate("2013/12/15 12:00"), _
                10000, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("17:30"))

FetchBars lContractSpec, lBarSpecFuture

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
Set lContractSpec = CreateContractSpecifier("ZZ3", "Z", "LIFFE", SecTypeFuture, "GBP", "201312")

Dim lBarSpecFuture As IFuture
Set lBarSpecFuture = CreateBarDataSpecifierFuture(FetchContract(lContractSpec, mContractStore), _
                GetTimePeriod(10, TimePeriodTickMovement), _
                CDate("2013/11/01 08:00"), _
                CDate("2013/12/15 12:00"), _
                500, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("17:30"))

FetchBars lContractSpec, lBarSpecFuture

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FetchYearsDataButton_Click()
Const ProcName As String = "FetchYearsDataButton_Click"
On Error GoTo Err

Dim lBarSpecFuture As IFuture
Dim lContractSpec As IContractSpecifier

Set lContractSpec = ContractSpecBuilder1.ContractSpecifier

Set lBarSpecFuture = CreateBarDataSpecifierFuture(FetchContract(lContractSpec, mContractStore), _
                GetTimePeriod(1, TimePeriodMinute), _
                CDate(Int(Now - 365#)), _
                Now, _
                365& * 1440&, _
                BarTypeTrade, _
                , _
                , _
                True)

FetchBars lContractSpec, lBarSpecFuture

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton1_Click()
Const ProcName As String = "ManyFetchButton1_Click"
On Error GoTo Err

fetchBarsForFTSEStock "AAL"
fetchBarsForFTSEStock "ABF"
fetchBarsForFTSEStock "ADM"
fetchBarsForFTSEStock "AGK"
fetchBarsForFTSEStock "AMEC"
fetchBarsForFTSEStock "ANTO"
fetchBarsForFTSEStock "ARM"
fetchBarsForFTSEStock "AU."
fetchBarsForFTSEStock "AV."
fetchBarsForFTSEStock "AZN"
' 10
fetchBarsForFTSEStock "BA."
fetchBarsForFTSEStock "BARC"
fetchBarsForFTSEStock "BATS"
fetchBarsForFTSEStock "BG."
fetchBarsForFTSEStock "BLND"
fetchBarsForFTSEStock "BLT"
fetchBarsForFTSEStock "BP."
fetchBarsForFTSEStock "BRBY"
fetchBarsForFTSEStock "BSY"
fetchBarsForFTSEStock "BT.A"
'20
fetchBarsForFTSEStock "CCL"
fetchBarsForFTSEStock "CNA"
fetchBarsForFTSEStock "CNE"
fetchBarsForFTSEStock "CPG"
fetchBarsForFTSEStock "CPI"
fetchBarsForFTSEStock "CSCG"
fetchBarsForFTSEStock "DGE"
fetchBarsForFTSEStock "EMG"
fetchBarsForFTSEStock "ENRC"
fetchBarsForFTSEStock "ESSR"
'30
fetchBarsForFTSEStock "EXPN"
fetchBarsForFTSEStock "FRES"
fetchBarsForFTSEStock "GFS"
fetchBarsForFTSEStock "GKN"
fetchBarsForFTSEStock "GSK"
fetchBarsForFTSEStock "HL."
fetchBarsForFTSEStock "HMSO"
fetchBarsForFTSEStock "HSBA"
fetchBarsForFTSEStock "IAG"
fetchBarsForFTSEStock "IAP"
'40
fetchBarsForFTSEStock "IHG"
fetchBarsForFTSEStock "III"
fetchBarsForFTSEStock "IMI"
fetchBarsForFTSEStock "IMT"
fetchBarsForFTSEStock "INVP"
fetchBarsForFTSEStock "IPR"
fetchBarsForFTSEStock "ISAT"
fetchBarsForFTSEStock "ISYS"
fetchBarsForFTSEStock "ITRK"
fetchBarsForFTSEStock "ITV"
'50

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ManyFetchButton2_Click()
Const ProcName As String = "ManyFetchButton2_Click"
On Error GoTo Err

fetchBarsForFTSEStock "JMAT"
fetchBarsForFTSEStock "KAZ"
fetchBarsForFTSEStock "KGF"
fetchBarsForFTSEStock "LAND"
fetchBarsForFTSEStock "LGEN"
fetchBarsForFTSEStock "LLOY"
fetchBarsForFTSEStock "LMI"
fetchBarsForFTSEStock "MKS"
fetchBarsForFTSEStock "MRW"
fetchBarsForFTSEStock "NG."
'60
fetchBarsForFTSEStock "NXT"
fetchBarsForFTSEStock "OML"
fetchBarsForFTSEStock "PFC"
fetchBarsForFTSEStock "PRU"
fetchBarsForFTSEStock "PSON"
fetchBarsForFTSEStock "RB."
fetchBarsForFTSEStock "RBS"
fetchBarsForFTSEStock "RDSA"
fetchBarsForFTSEStock "RDSB"
fetchBarsForFTSEStock "REL"
'70
fetchBarsForFTSEStock "REX"
fetchBarsForFTSEStock "RIO"
fetchBarsForFTSEStock "RR."
fetchBarsForFTSEStock "RRS"
fetchBarsForFTSEStock "RSA"
fetchBarsForFTSEStock "RSL"
fetchBarsForFTSEStock "SAB"
fetchBarsForFTSEStock "SBRY"
fetchBarsForFTSEStock "SDR"
fetchBarsForFTSEStock "SDRC"
'80
fetchBarsForFTSEStock "SGE"
fetchBarsForFTSEStock "SHP"
fetchBarsForFTSEStock "SL."
fetchBarsForFTSEStock "SMIN"
fetchBarsForFTSEStock "SN."
fetchBarsForFTSEStock "SRP"
fetchBarsForFTSEStock "SSE"
fetchBarsForFTSEStock "STAN"
fetchBarsForFTSEStock "SVT"
fetchBarsForFTSEStock "TLW"
'90
fetchBarsForFTSEStock "TSCO"
fetchBarsForFTSEStock "TT."
fetchBarsForFTSEStock "ULVR"
fetchBarsForFTSEStock "UU."
fetchBarsForFTSEStock "VED"
fetchBarsForFTSEStock "VOD"
fetchBarsForFTSEStock "WEIR"
fetchBarsForFTSEStock "WG."
fetchBarsForFTSEStock "WOS"
fetchBarsForFTSEStock "WPP"
'100
fetchBarsForFTSEStock "WTB"
fetchBarsForFTSEStock "XTA"

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

Private Sub FetchBars(ByVal pContractSpec As IContractSpecifier, ByVal pBarSpecFuture As IFuture)
Const ProcName As String = "FetchBars"
On Error GoTo Err

Dim lListener As New BarListener
lListener.Initialise pBarSpecFuture, mHistDataStore, pContractSpec, (ShowBarOnReceipt.Value = vbChecked), (ShowBarsAtEndCheck.Value = vbChecked)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fetchBarsForFTSEStock(ByVal pSymbol As String)
Const ProcName As String = "fetchBarsForFTSEStock"
On Error GoTo Err

Dim lBarSpecFuture As IFuture
Dim lContractSpec As IContractSpecifier

Set lContractSpec = CreateContractSpecifier(pSymbol, , "LSE", SecTypeStock, "GBP")

Set lBarSpecFuture = CreateBarDataSpecifierFuture(FetchContract(lContractSpec, mContractStore), _
                GetTimePeriod(5, TimePeriodMinute), _
                Now - 7#, _
                Now, _
                2000, _
                BarTypeTrade, _
                , _
                , _
                False, _
                CDate("08:00"), _
                CDate("16:30"))

FetchBars lContractSpec, lBarSpecFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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


