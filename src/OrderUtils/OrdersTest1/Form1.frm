VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#338.0#0"; "TradingUI27.ocx"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   12615
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   16695
   LinkTopic       =   "Form1"
   ScaleHeight     =   12615
   ScaleWidth      =   16695
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWButton ClosePositionsButton 
      Height          =   1215
      Left            =   14640
      TabIndex        =   7
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
      Caption         =   "Close All Positions"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
   Begin VB.CheckBox ShowMarketDepthCheck 
      Caption         =   "Show market depth"
      Height          =   375
      Left            =   11760
      TabIndex        =   6
      Top             =   1560
      Width           =   3015
   End
   Begin TradingUI27.DOMDisplay DOMDisplay 
      Height          =   10455
      Left            =   11760
      TabIndex        =   5
      Top             =   2040
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   18441
   End
   Begin TradingUI27.OrdersSummary OrdersSummary 
      Height          =   3495
      Left            =   120
      TabIndex        =   3
      Top             =   6240
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   6165
   End
   Begin TradingUI27.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   3510
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6191
      ForeColor       =   -2147483640
      ModeAdvanced    =   -1  'True
   End
   Begin VB.TextBox LogText 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   9840
      Width           =   11535
   End
   Begin TradingUI27.OrderTicket OrderTicket 
      Height          =   6135
      Left            =   2880
      TabIndex        =   2
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   10821
   End
   Begin VB.CheckBox UsePositionManagerCheck 
      Caption         =   "Use Position Manager"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   1935
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

Implements IChangeListener
Implements IErrorListener
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

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mClientId                                   As Long

Private mClient                                     As Client

Private mContractStore                              As IContractStore

Private mOrderManager                               As New OrderManager

Private mPositionManagersLive                       As PositionManagers
Private mPositionManagersSimulated                  As PositionManagers

Private mPositionManagerLive                        As PositionManager
Private mPositionManagerSimulated                   As PositionManager

Private mOrderContextsLive                          As OrderContexts
Private mOrderContextsSimulated                     As OrderContexts

Private mOrderContextLive                           As OrderContext
Private mOrderContextSimulated                      As OrderContext

Private mOrderSubmitterLive                         As IOrderSubmitter
Private mOrderSubmitterFactorySimulated             As IOrderSubmitterFactory

Private mContractFuture                             As IFuture
Private mContract                                   As IContract

Private mMarketDataManager                          As IMarketDataManager

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

Set mUnhandledErrorHandler = UnhandledErrorHandler

InitialiseCommonControls

InitialiseTWUtilities

ApplicationGroupName = "TradeWright"
ApplicationName = "OrdersTest1"
DefaultLogLevel = LogLevelHighDetail
SetupDefaultLogging Command
GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box

Set mFutureWaiter = New FutureWaiter

Set mPositionManagersLive = mOrderManager.PositionManagersLive
Set mPositionManagersSimulated = mOrderManager.PositionManagersSimulated

mClientId = 1132256741
Set mClient = GetClient("Essy", 7497, mClientId, , , True, , Me)

Set mContractStore = mClient.GetContractStore
Set mOrderSubmitterLive = mClient.CreateOrderSubmitter
mOrderSubmitterLive.AddOrderSubmissionListener Me
Set mMarketDataManager = CreateRealtimeDataManager(mClient.GetMarketDataFactory)

Set mOrderSubmitterFactorySimulated = New SimOrderSubmitterFactory

OrdersSummary.Initialise mMarketDataManager
OrdersSummary.MonitorPositions mPositionManagersLive
OrdersSummary.MonitorPositions mPositionManagersSimulated

mFutureWaiter.Add CreateFutureFromTask(mOrderManager.RecoverOrdersFromPreviousSession(ApplicationName, _
                                        CreateOrderPersistenceDataStore(ApplicationSettingsFolder), _
                                        mClient, _
                                        mMarketDataManager, _
                                        mClient))
                                        
Set mTheme = New BlackTheme
Me.BackColor = mTheme.BaseColor
gApplyTheme mTheme, Me.Controls

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
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
On Error GoTo Err

Dim lChangeType As PositionManagerChangeTypes
Dim lPositionManager As PositionManager

lChangeType = ev.ChangeType
Set lPositionManager = ev.Source

Select Case lChangeType
Case PositionSizeChanged
    ensurePositionFormVisible lPositionManager
Case ProviderReadinessChanged

Case PositionClosed

End Select


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IErrorListener Interface Members
'@================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
Const ProcName As String = "IErrorListener_Notify"
On Error GoTo Err

LogMessage "Error " & ev.ErrorCode & ": " & ev.ErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Sub IOrderSubmissionListener_NotifyMessage(ByVal pOrderId As String, ByVal pMessage As String)
LogMessage "Message: " & pMessage
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

Private Sub ContractSpecBuilder1_NotReady()
Const ProcName As String = "ContractSpecBuilder1_NotReady"
On Error GoTo Err

If Not mContractFuture Is Nothing Then
    If mContractFuture.IsPending Then
        mContractFuture.Cancel
    End If
    Set mContractFuture = Nothing
End If

OrderTicket.Clear
DOMDisplay.Finish
releaseUnusedPositionManager mPositionManagerLive
releaseUnusedPositionManager mPositionManagerSimulated
Set mOrderContextsLive = Nothing
Set mOrderContextsSimulated = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ContractSpecBuilder1_Ready()
Const ProcName As String = "ContractSpecBuilder1_Ready"
On Error GoTo Err

If Not mContractFuture Is Nothing Then
    If mContractFuture.IsPending Then
        mContractFuture.Cancel
    End If
    Set mContractFuture = Nothing
End If
Set mContract = Nothing

OrderTicket.Clear
DOMDisplay.Finish
releaseUnusedPositionManager mPositionManagerLive
releaseUnusedPositionManager mPositionManagerSimulated
Set mOrderContextsLive = Nothing
Set mOrderContextsSimulated = Nothing

Set mContractFuture = FetchContract(ContractSpecBuilder1.ContractSpecifier, mContractStore)
mFutureWaiter.Add mContractFuture

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OrdersSummary_Click()
Const ProcName As String = "OrdersSummary_Click"
On Error GoTo Err

If Not OrdersSummary.SelectedItem Is Nothing Then OrderTicket.ShowBracketOrder OrdersSummary.SelectedItem, OrdersSummary.SelectedOrderRole

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OrderTicket_NeedLiveOrderContext()
Const ProcName As String = "OrderTicket_NeedLiveOrderContext"
On Error GoTo Err

OrderTicket.SetLiveOrderContext mOrderContextLive

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OrderTicket_NeedSimulatedOrderContext()
Const ProcName As String = "OrderTicket_NeedSimulatedOrderContext"
On Error GoTo Err

OrderTicket.SetSimulatedOrderContext mOrderContextSimulated

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ShowMarketDepthCheck_Click()
Const ProcName As String = "ShowMarketDepthCheck_Click"
On Error GoTo Err

If ShowMarketDepthCheck.Value = vbChecked Then
    If Not mPositionManagerLive Is Nothing Then DOMDisplay.DataSource = mPositionManagerLive.DataSource
Else
    DOMDisplay.Finish
End If

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

If ev.Future.IsAvailable Then
    If TypeOf ev.Future.Value Is PositionManagers Then
        Dim lPM As PositionManager
        For Each lPM In mPositionManagersLive
            If lPM.PositionSize <> 0 Or lPM.PendingPositionSize <> 0 Then ensurePositionFormVisible lPM
        Next
    ElseIf TypeOf ev.Future.Value Is IContract Then
        Set mContract = ev.Future.Value
        If mContract.Specifier.SecType = SecTypeCombo Or _
            mContract.Specifier.SecType = SecTypeIndex _
        Then
            LogMessage "Non-tradeable contract found"
        ElseIf IsContractExpired(mContract.ExpiryDate) Then
            LogMessage "Expired contract found"
        Else
            LogMessage "Found contract " & mContract.Specifier.ToString
            setupOrderContext mContract, UsePositionManagerCheck.Value
        End If
    End If
ElseIf ev.Future.IsFaulted Then
    LogMessage ev.Future.ErrorMessage
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

Private Sub ensurePositionFormVisible(ByVal pPositionManager As PositionManager)
Const ProcName As String = "ensurePositionFormVisible"
On Error GoTo Err

Static sPositionManagerForms As New EnumerableCollection
Dim lForm As PositionForm

If UsePositionManagerCheck.Value = vbUnchecked Then Exit Sub

If Not sPositionManagerForms.Contains(pPositionManager.Name) Then
    Set lForm = New PositionForm
    lForm.Initialise pPositionManager
    sPositionManagerForms.Add lForm, pPositionManager.Name
Else
    Set lForm = sPositionManagerForms.Item(pPositionManager.Name)
End If

lForm.Show vbModeless, Me

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

Private Function getDataSource(pContract) As IMarketDataSource
Const ProcName As String = "getDataSource"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = mMarketDataManager.CreateMarketDataSource(CreateFuture(pContract), False)
lDataSource.StartMarketData
Set getDataSource = lDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDouble(ByVal pText As String) As Double
If Len(pText) = 0 Then
    getDouble = 0
Else
    getDouble = CDbl(pText)
End If
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

If Not mClient Is Nothing Then mClient.Finish

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

Private Sub releaseUnusedPositionManager(ByVal pPositionManager As PositionManager)
Const ProcName As String = "releaseUnusedPositionManager"
On Error GoTo Err

If pPositionManager Is Nothing Then Exit Sub
If pPositionManager.IsFinished Then Exit Sub
If pPositionManager.PositionSize <> 0 Or pPositionManager.PendingPositionSize <> 0 Then Exit Sub

pPositionManager.RemoveChangeListener Me
pPositionManager.DataSource.StopMarketData

If pPositionManager.IsSimulated Then
    mPositionManagersSimulated.Remove pPositionManager
Else
    mPositionManagersLive.Remove pPositionManager
End If

pPositionManager.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderContext( _
                ByVal pContract As IContract, _
                ByVal pUsePositionManager As Boolean)
Const ProcName As String = "setupOrderContext"
On Error GoTo Err

If pUsePositionManager Then
    'releaseUnusedPositionManager mPositionManagerLive
    Set mPositionManagerLive = Nothing
    'releaseUnusedPositionManager mPositionManagerSimulated
    Set mPositionManagerSimulated = Nothing
    
    Dim lDataSource As IMarketDataSource
    If mPositionManagersLive.Contains(pContract.Specifier.Key) Then
        Set mPositionManagerLive = mPositionManagersLive.Item(pContract.Specifier.Key)
        Set lDataSource = mPositionManagerLive.DataSource
    End If
    
    If mPositionManagersSimulated.Contains(pContract.Specifier.Key) Then
        Set mPositionManagerSimulated = mPositionManagersSimulated.Item(pContract.Specifier.Key)
        Set lDataSource = mPositionManagerSimulated.DataSource
    End If
    
    If lDataSource Is Nothing Then
        Set lDataSource = getDataSource(pContract)
        If ShowMarketDepthCheck.Value = vbChecked Then DOMDisplay.DataSource = lDataSource
    End If
    
    If mPositionManagerLive Is Nothing Then
        Set mPositionManagerLive = mOrderManager.CreateRecoverablePositionManager(pContract.Specifier.Key, lDataSource, mClient, ApplicationName, "DefaultGroup")
        mPositionManagerLive.AddChangeListener Me
        mPositionManagerLive.OrderSubmitter.AddOrderSubmissionListener Me
    End If
    
    If mPositionManagerSimulated Is Nothing Then
        Set mPositionManagerSimulated = mOrderManager.CreatePositionManager(pContract.Specifier.Key, lDataSource, mOrderSubmitterFactorySimulated, "DefaultGroup", True)
        mPositionManagerSimulated.AddChangeListener Me
        mPositionManagerSimulated.OrderSubmitter.AddOrderSubmissionListener Me
    End If
    
    Set mOrderContextsLive = mPositionManagerLive.OrderContexts
    Set mOrderContextsSimulated = mPositionManagerSimulated.OrderContexts
Else
    Set mOrderContextsLive = mOrderManager.GetOrderContexts(pContract.Specifier.Key, False)
    If mOrderContextsLive Is Nothing Then Set mOrderContextsLive = mOrderManager.CreateOrderContexts(pContract.Specifier.Key, CreateFuture(pContract), mOrderSubmitterLive)
    mOrderContextsLive.OrderSubmitter.AddOrderSubmissionListener Me

    Set mOrderContextsSimulated = mOrderManager.GetOrderContexts(pContract.Specifier.Key, True)
    If mOrderContextsSimulated Is Nothing Then Set mOrderContextsSimulated = mOrderManager.CreateOrderContexts(pContract.Specifier.Key, CreateFuture(pContract), mOrderSubmitterFactorySimulated.CreateOrderSubmitter, , , True)
    mOrderContextsSimulated.OrderSubmitter.AddOrderSubmissionListener Me
End If

Set mOrderContextLive = mOrderContextsLive.DefaultOrderContext
Set mOrderContextSimulated = mOrderContextsSimulated.DefaultOrderContext

OrderTicket.SetMode OrderTicketModes.OrderTicketModeLiveAndSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



