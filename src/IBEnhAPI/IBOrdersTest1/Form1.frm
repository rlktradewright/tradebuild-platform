VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#292.0#0"; "TradingUI27.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11475
   LinkTopic       =   "Form1"
   ScaleHeight     =   7800
   ScaleWidth      =   11475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Height          =   255
      Left            =   7440
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox ClientIdText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Text            =   "1143256749"
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox PortText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3480
      TabIndex        =   1
      Text            =   "7497"
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox ServerText 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
   Begin VB.TextBox ContractText 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   1560
      Width           =   8175
   End
   Begin VB.CommandButton BuyBracketButton 
      Caption         =   "Buy"
      Height          =   495
      Left            =   10440
      TabIndex        =   26
      Top             =   2400
      Width           =   855
   End
   Begin VB.CommandButton SellBracketButton 
      Caption         =   "Sell"
      Height          =   495
      Left            =   10440
      TabIndex        =   25
      Top             =   3000
      Width           =   855
   End
   Begin VB.TextBox TargetTriggerPriceText 
      Height          =   285
      Left            =   9120
      TabIndex        =   23
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox TargetPriceText 
      Height          =   285
      Left            =   8040
      TabIndex        =   22
      Top             =   3120
      Width           =   975
   End
   Begin VB.ComboBox TargetOrderTypeCombo 
      Height          =   315
      Left            =   6480
      TabIndex        =   21
      Top             =   3120
      Width           =   1455
   End
   Begin VB.TextBox StopLossTriggerPriceText 
      Height          =   285
      Left            =   9120
      TabIndex        =   19
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox StopLossPriceText 
      Height          =   285
      Left            =   8040
      TabIndex        =   18
      Top             =   2760
      Width           =   975
   End
   Begin VB.ComboBox StopLossOrderTypeCombo 
      Height          =   315
      Left            =   6480
      TabIndex        =   17
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox EntryTriggerPriceText 
      Height          =   285
      Left            =   9120
      TabIndex        =   13
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox EntryPriceText 
      Height          =   285
      Left            =   8040
      TabIndex        =   12
      Top             =   2400
      Width           =   975
   End
   Begin VB.ComboBox EntryOrderTypeCombo 
      Height          =   315
      Left            =   6480
      TabIndex        =   10
      Top             =   2400
      Width           =   1455
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3690
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   5556
      ForeColor       =   -2147483640
      ModeAdvanced    =   -1  'True
   End
   Begin VB.CommandButton SellButton 
      Caption         =   "Sell"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton BuyButton 
      Caption         =   "Buy"
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox QuantityText 
      Alignment       =   1  'Right Justify
      Height          =   330
      Left            =   3120
      TabIndex        =   5
      Text            =   "1"
      Top             =   2160
      Width           =   855
   End
   Begin VB.TextBox LogText 
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4800
      Width           =   11295
   End
   Begin VB.Label ClientIdLabel 
      Caption         =   "Client id"
      Height          =   375
      Left            =   4800
      TabIndex        =   30
      Top             =   120
      Width           =   855
   End
   Begin VB.Label PortLabel 
      Caption         =   "Port"
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label 
      Caption         =   "Server"
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   28
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Target"
      Height          =   375
      Left            =   5640
      TabIndex        =   24
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label8 
      Caption         =   "Stop loss"
      Height          =   375
      Left            =   5640
      TabIndex        =   20
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Entry"
      Height          =   375
      Left            =   5640
      TabIndex        =   16
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Trigger price"
      Height          =   375
      Left            =   9120
      TabIndex        =   15
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Price"
      Height          =   375
      Left            =   8040
      TabIndex        =   14
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Order type"
      Height          =   375
      Left            =   6480
      TabIndex        =   11
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label 
      Caption         =   "Qty"
      Height          =   255
      Index           =   0
      Left            =   2640
      TabIndex        =   9
      Top             =   2160
      Width           =   375
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

Implements IOrderSubmissionListener
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

Private Const CaptionConnect                        As String = "Connect"
Private Const CaptionDisconnect                     As String = "Disconnect"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mClientId                                   As Long

Private mClient                                     As Client

Private mOrderSubmitter                             As IOrderSubmitter
Private mContractStore                              As IContractStore

Private mContractFuture                             As IFuture
Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContract                                   As IContract

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
Set mUnhandledErrorHandler = UnhandledErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "IBOrdersTest127"
DefaultLogLevel = LogLevelHighDetail
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

setupEntryOrderTypeCombo
setupStopLossOrderTypeCombo
setupTargetOrderTypeCombo

ConnectButton.Caption = CaptionConnect

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
mClient.Finish
End Sub

'@================================================================================
' IOrderSubmissionListener Interface Members
'@================================================================================

Private Sub IOrderSubmissionListener_NotifyError(ByVal pOrderId As String, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
LogMessage "Error " & pErrorCode & ": " & pErrorMsg
End Sub

Private Sub IOrderSubmissionListener_NotifyExecutionReport(ByVal pExecutionReport As IExecutionReport)
LogMessage "Execution: " & pExecutionReport.BrokerId & "/" & pExecutionReport.Id & _
            "; Qty=" & pExecutionReport.Quantity & _
            "; Price=" & pExecutionReport.Price
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderReport(ByVal pOrderReport As IOrderReport)
LogMessage "Order details: " & pOrderReport.BrokerId & "/" & pOrderReport.Id
End Sub

Private Sub IOrderSubmissionListener_NotifyOrderStatusReport(ByVal pOrderStatusReport As IOrderStatusReport)
LogMessage "Order status: " & pOrderStatusReport.BrokerId & "/" & pOrderStatusReport.OrderId & _
            "; Status=" & OrderStatusToString(pOrderStatusReport.Status)
End Sub

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    LogMessage "Disconnected from TWS: " & pMessage
Case ApiConnConnecting
    LogMessage "Connecting to TWS: " & pMessage
Case ApiConnConnected
    mClient.SetTwsLogLevel TwsLogLevelDetail
    LogMessage "Connected to TWS: " & pMessage
    ConnectButton.Caption = CaptionDisconnect
    ConnectButton.Enabled = True
Case ApiConnFailed
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

Private Sub BuyBracketButton_Click()
Const ProcName As String = "BuyBracketButton_Click"
On Error GoTo Err

submitBracketOrder mContract, OrderActionBuy

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub BuyButton_Click()
Const ProcName As String = "BuyButton_Click"
On Error GoTo Err

Dim lOrder As New Order

lOrder.ContractSpecifier = ContractSpecBuilder1.ContractSpecifier
lOrder.Action = OrderActionBuy
lOrder.Quantity = CLng(QuantityText)
lOrder.OrderType = OrderTypeMarket
lOrder.ETradeOnly = True
lOrder.FirmQuoteOnly = True

mOrderSubmitter.PlaceOrder lOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ConnectButton_Click()
If ConnectButton.Caption = CaptionConnect Then
    ConnectButton.Enabled = False
    
    mClientId = CLng(ClientIdText.Text)
    Set mClient = GetClient(ServerText.Text, CLng(PortText.Text), mClientId, , , , Me)
    
    Set mOrderSubmitter = mClient.CreateOrderSubmitter
    mOrderSubmitter.AddOrderSubmissionListener Me
    Set mContractStore = mClient.GetContractStore
    
    Set mFutureWaiter = New FutureWaiter
Else
    mClient.Finish
    ConnectButton.Caption = CaptionConnect
End If
End Sub

Private Sub ContractSpecBuilder1_NotReady()
BuyButton.Enabled = False
SellButton.Enabled = False
BuyBracketButton.Enabled = False
SellBracketButton.Enabled = False
ContractText.Text = ""
End Sub

Private Sub ContractSpecBuilder1_Ready()
BuyButton.Enabled = False
SellButton.Enabled = False
BuyBracketButton.Enabled = False
SellBracketButton.Enabled = False
ContractText.Text = ""

If Not mContractFuture Is Nothing Then mContractFuture.Cancel
Set mContractFuture = FetchContract(ContractSpecBuilder1.ContractSpecifier, mContractStore)
mFutureWaiter.Add mContractFuture
End Sub

Private Sub SellBracketButton_Click()
Const ProcName As String = "SellBracketButton_Click"
On Error GoTo Err

submitBracketOrder mContract, OrderActionSell

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub SellButton_Click()
Const ProcName As String = "SellButton_Click"
On Error GoTo Err

Dim lOrder As New Order

lOrder.ContractSpecifier = ContractSpecBuilder1.ContractSpecifier
lOrder.Action = OrderActionSell
lOrder.Quantity = CLng(QuantityText)
lOrder.OrderType = OrderTypeMarket
lOrder.ETradeOnly = True
lOrder.FirmQuoteOnly = True

mOrderSubmitter.PlaceOrder lOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsPending Then Exit Sub

BuyButton.Enabled = False
SellButton.Enabled = False
BuyBracketButton.Enabled = False
SellBracketButton.Enabled = False

If ev.Future.IsCancelled Then
    'LogMessage "Contract future creation cancelled"
ElseIf ev.Future.IsFaulted Then
    LogMessage ev.Future.ErrorMessage
Else
    Set mContract = ev.Future.Value
    ContractText.Text = mContract.Specifier.LocalSymbol & " (" & mContract.Specifier.Exchange & ")  " & mContract.Description
    BuyButton.Enabled = True
    SellButton.Enabled = True
    BuyBracketButton.Enabled = True
    SellBracketButton.Enabled = True
End If

Set mContractFuture = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

mClient.Finish

handleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
'UnhandledErrorHandler.Handled = True
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

Private Sub addItemToCombo( _
                ByVal pCombo As ComboBox, _
                ByVal pItemText As String, _
                ByVal pItemData As Long)
Const ProcName As String = "addItemToCombo"
Dim failpoint As String
On Error GoTo Err

pCombo.AddItem pItemText
pCombo.ItemData(pCombo.ListCount - 1) = pItemData

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

Private Sub setupEntryOrderTypeCombo()
addItemToCombo EntryOrderTypeCombo, BracketEntryTypeToString(BracketEntryTypeMarket), BracketEntryTypeMarket
addItemToCombo EntryOrderTypeCombo, BracketEntryTypeToString(BracketEntryTypeLimit), BracketEntryTypeLimit
addItemToCombo EntryOrderTypeCombo, BracketEntryTypeToString(BracketEntryTypeStop), BracketEntryTypeStop
addItemToCombo EntryOrderTypeCombo, BracketEntryTypeToString(BracketEntryTypeStopLimit), BracketEntryTypeStopLimit
EntryOrderTypeCombo.ListIndex = 0
End Sub

Private Sub setupStopLossOrderTypeCombo()
addItemToCombo StopLossOrderTypeCombo, BracketStopLossTypeToString(BracketStopLossTypeStop), BracketStopLossTypeStop
addItemToCombo StopLossOrderTypeCombo, BracketStopLossTypeToString(BracketStopLossTypeStopLimit), BracketStopLossTypeStopLimit
StopLossOrderTypeCombo.ListIndex = 0
End Sub

Private Sub setupTargetOrderTypeCombo()
addItemToCombo TargetOrderTypeCombo, BracketTargetTypeToString(BracketTargetTypeLimit), BracketTargetTypeLimit
End Sub

Private Sub submitBracketOrder(ByVal pContract As IContract, ByVal pAction As OrderActions)
Const ProcName As String = "submitBracketOrder"
On Error GoTo Err

Dim lBracketOrder As New BracketOrder
lBracketOrder.Contract = pContract

Dim lEntryOrder As New Order
lEntryOrder.OrderType = BracketEntryTypeToOrderType(EntryOrderTypeCombo.ItemData(EntryOrderTypeCombo.ListIndex))
lEntryOrder.Action = pAction
lEntryOrder.Quantity = CLng(QuantityText)
lEntryOrder.ETradeOnly = True
lEntryOrder.FirmQuoteOnly = True
If EntryPriceText.Text <> "" Then lEntryOrder.LimitPrice = CDbl(EntryPriceText.Text)
If EntryTriggerPriceText.Text <> "" Then lEntryOrder.TriggerPrice = CDbl(EntryTriggerPriceText.Text)

lBracketOrder.EntryOrder = lEntryOrder

If StopLossOrderTypeCombo.Text <> "" Then
    Dim lStopLossOrder As New Order
    lStopLossOrder.OrderType = BracketStopLossTypeToOrderType(StopLossOrderTypeCombo.ItemData(StopLossOrderTypeCombo.ListIndex))
    lStopLossOrder.Action = IIf(pAction = OrderActionBuy, OrderActionSell, OrderActionBuy)
    lStopLossOrder.Quantity = CLng(QuantityText)
    lStopLossOrder.ETradeOnly = True
    lStopLossOrder.FirmQuoteOnly = True
    If StopLossPriceText.Text <> "" Then lStopLossOrder.LimitPrice = CDbl(StopLossPriceText.Text)
    If StopLossTriggerPriceText.Text <> "" Then lStopLossOrder.TriggerPrice = CDbl(StopLossTriggerPriceText.Text)
    
    lBracketOrder.StopLossOrder = lStopLossOrder
End If

If TargetOrderTypeCombo.Text <> "" Then
    Dim lTargetOrder As New Order
    lTargetOrder.OrderType = BracketTargetTypeToOrderType(TargetOrderTypeCombo.ItemData(TargetOrderTypeCombo.ListIndex))
    lTargetOrder.Action = IIf(pAction = OrderActionBuy, OrderActionSell, OrderActionBuy)
    lTargetOrder.Quantity = CLng(QuantityText)
    lTargetOrder.ETradeOnly = True
    lTargetOrder.FirmQuoteOnly = True
    If TargetPriceText.Text <> "" Then lTargetOrder.LimitPrice = CDbl(TargetPriceText.Text)
    If TargetTriggerPriceText.Text <> "" Then lTargetOrder.TriggerPrice = CDbl(TargetTriggerPriceText.Text)
    
    lBracketOrder.TargetOrder = lTargetOrder
End If

mOrderSubmitter.ExecuteBracketOrder lBracketOrder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub



