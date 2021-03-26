VERSION 5.00
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#376.0#0"; "TradingUI27.ocx"
Begin VB.Form Form1 
   Caption         =   "Ticker Grid Test1"
   ClientHeight    =   10065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14415
   LinkTopic       =   "Form1"
   ScaleHeight     =   10065
   ScaleWidth      =   14415
   StartUpPosition =   3  'Windows Default
   Begin TradingUI27.TickerGrid TickerGrid 
      Height          =   7095
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12515
      AllowUserReordering=   3
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox LogText 
      Height          =   2655
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7320
      Width           =   14175
   End
   Begin TradingUI27.ContractSearch ContractSearch 
      Height          =   7095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12568
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

Implements IDeferredAction
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

Private Enum DeferredActions
    DeferredActionSaveConfig = 1
End Enum

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

Private mDataClient                                 As Client
Private mContractClient                             As Client

Private mTickersStarted                             As Boolean

Private mMarketDataManager                          As IMarketDataManager
Private mContractStore                              As IContractStore

Private mNoLogfile                                  As Boolean

Private WithEvents mContractSelectionHelper         As ContractSelectionHelper
Attribute mContractSelectionHelper.VB_VarHelpID = -1

Private mPreferredGridRow                           As Long

Private WithEvents mConfigStore                     As ConfigurationStore
Attribute mConfigStore.VB_VarHelpID = -1
Private mMarketDataManagerConfig                    As ConfigurationSection
Private mTickerGridConfig                           As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

Set mUnhandledErrorHandler = UnhandledErrorHandler

InitialiseTWUtilities

ApplicationGroupName = "TradeWright"
ApplicationName = "TickerGridTest1"
SetupDefaultLogging Command

Set mConfigStore = getConfigStore
Set mMarketDataManagerConfig = mConfigStore.AddPrivateConfigurationSection("/MarketDataManager")
Set mTickerGridConfig = mConfigStore.AddPrivateConfigurationSection("/TickerGrid")
If mConfigStore.Dirty Then mConfigStore.Save

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrSecurityException Then
    mNoLogfile = True
    Resume Next
End If
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

GetLogger("log").AddLogListener Me  ' so that log entries of infotype 'log' will be written to the logging text box
If mNoLogfile Then
    LogMessage "Can't open the log file: " & DefaultLogFileName(Command)
Else
    LogMessage "Logging to file: " & DefaultLogFileName(Command)
End If

mClientId = 1132256741
Set mDataClient = GetClient("Essy", 7497, mClientId, , , ApiMessageLoggingOptionDefault, ApiMessageLoggingOptionNone, False, , Me)
mDataClient.SetTwsLogLevel TwsLogLevelDetail

Set mContractClient = GetClient("Essy", 7497, mClientId + 1, , , ApiMessageLoggingOptionDefault, ApiMessageLoggingOptionNone, False, , Me)

Set mContractStore = mContractClient.GetContractStore
Set mMarketDataManager = CreateRealtimeDataManager(mDataClient.GetMarketDataFactory, mDataClient.GetContractStore)

ContractSearch.Initialise mContractStore, Nothing
ContractSearch.IncludeHistoricalContracts = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
If Not mDataClient Is Nothing Then mDataClient.Finish
If Not mContractClient Is Nothing Then mContractClient.Finish
TerminateTWUtilities
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Select Case CLng(Data)
Case DeferredActions.DeferredActionSaveConfig
    If mConfigStore.Dirty Then mConfigStore.Save
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    LogMessage "Connected to TWS: " & pMessage
    
    If Not mTickersStarted Then
        mTickersStarted = True
        mMarketDataManager.LoadFromConfig mMarketDataManagerConfig
        TickerGrid.Initialise mMarketDataManager, mTickerGridConfig
    End If
Case ApiConnFailed
    LogMessage "Failed to connect to TWS: " & pMessage
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ContractSearch_Action()
Const ProcName As String = "ContractSearch_Action"
On Error GoTo Err

Dim lContract As IContract

For Each lContract In ContractSearch.SelectedContracts
    TickerGrid.StartTickerFromContract lContract
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_Error(ev As ErrorEventData)
Const ProcName As String = "TickerGrid_Error"
On Error GoTo Err

LogMessage "Error: " & ev.ErrorMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_TickerSymbolEntered(ByVal pSymbol As String, ByVal pPreferredRow As Long)
Const ProcName As String = "TickerGrid_TickerSymbolEntered"
On Error GoTo Err

mPreferredGridRow = pPreferredRow
Set mContractSelectionHelper = CreateContractSelectionHelper( _
                                        CreateContractSpecifier(Symbol:=pSymbol), _
                                        mContractStore)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mConfigStore Event Handlers
'@================================================================================

Private Sub mConfigStore_Change(ev As ChangeEventData)
Const ProcName As String = "mConfigStore_Change"
On Error GoTo Err

If ev.ChangeType = ConfigChangeTypes.ConfigDirty Then DeferAction Me, DeferredActions.DeferredActionSaveConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mContractSelectionHelper Event Handlers
'@================================================================================

Private Sub mContractSelectionHelper_Cancelled()
Const ProcName As String = "mContractSelectionHelper_Cancelled"
On Error GoTo Err

LogMessage "Contract search cancelled"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mContractSelectionHelper_Error(ev As ErrorEventData)
Const ProcName As String = "mContractSelectionHelper_Error"
On Error GoTo Err

Err.Raise ev.ErrorCode, ev.Source, ev.ErrorMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mContractSelectionHelper_Ready()
Const ProcName As String = "mContractSelectionHelper_Ready"
On Error GoTo Err

If mContractSelectionHelper.Contracts.Count = 0 Then
    LogMessage "Invalid symbol"
Else
    TickerGrid.StartTickerFromContract mContractSelectionHelper.Contracts.ItemAtIndex(1), mPreferredGridRow
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mContractSelectionHelper_ShowContractSelector()
Const ProcName As String = "mContractSelectionHelper_ShowContractSelector"
On Error GoTo Err

Dim f As fContractSelector

Set f = New fContractSelector
f.Initialise mContractSelectionHelper.Contracts, mContractStore, True
f.Show vbModal, Me

Dim lContracts As IContracts
Set lContracts = f.SelectedContracts
If lContracts.Count = 0 Then Exit Sub

Dim lContract As IContract
For Each lContract In lContracts
    TickerGrid.StartTickerFromContract lContract, mPreferredGridRow
    mPreferredGridRow = mPreferredGridRow + 1
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

If Not mDataClient Is Nothing Then mDataClient.Finish
If Not mContractClient Is Nothing Then mContractClient.Finish

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

Private Function formatLogRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "formatLogRecord"
Static formatter As ILogFormatter

On Error GoTo Err

If formatter Is Nothing Then Set formatter = CreateBasicLogFormatter(TimestampFormats.TimestampTimeOnlyLocal)
formatLogRecord = formatter.FormatRecord(Logrec)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getConfigStore() As ConfigurationStore
Const ProcName As String = "getConfigStore"
On Error GoTo Err

Set getConfigStore = GetDefaultConfigurationStore(Command, "1.0", False, ConfigFileOptionFirstArg)
If getConfigStore Is Nothing Then Set getConfigStore = GetDefaultConfigurationStore(Command, "1.0", True, ConfigFileOptionFirstArg)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

If Not mDataClient Is Nothing Then mDataClient.Finish
If Not mContractClient Is Nothing Then mContractClient.Finish

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





