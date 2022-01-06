VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Object = "{6C945B95-5FA7-4850-AAF3-2D2AA0476EE1}#375.0#0"; "TradingUI27.ocx"
Begin VB.Form fTradeSkilDemo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TradeSkil Demo Edition"
   ClientHeight    =   9960
   ClientLeft      =   225
   ClientTop       =   345
   ClientWidth     =   16665
   Icon            =   "fTradeSkilDemo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9960
   ScaleWidth      =   16665
   Begin TradeSkilDemo27.InfoPanel InfoPanel 
      Height          =   4755
      Left            =   4320
      TabIndex        =   5
      Top             =   4800
      Visible         =   0   'False
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8387
   End
   Begin VB.PictureBox ShowInfoPanelPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   16320
      MouseIcon       =   "fTradeSkilDemo.frx":3307A
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":331CC
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   0
      ToolTipText     =   "Show Information Panel"
      Top             =   9345
      Width           =   240
   End
   Begin VB.PictureBox ShowFeaturesPanelPicture 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   0
      MouseIcon       =   "fTradeSkilDemo.frx":33756
      MousePointer    =   99  'Custom
      Picture         =   "fTradeSkilDemo.frx":338A8
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      ToolTipText     =   "Show Features Panel"
      Top             =   120
      Width           =   240
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   9585
      Width           =   16665
      _ExtentX        =   29395
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   23733
            Key             =   "status"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   0
            Key             =   "timezone"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   2
            Bevel           =   0
            Key             =   "datetime"
         EndProperty
      EndProperty
   End
   Begin TradingUI27.TickerGrid TickerGrid1 
      Height          =   4695
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   8281
      AllowUserReordering=   3
      BackColorFixed  =   16053492
      RowBackColorOdd =   16316664
      RowBackColorEven=   15658734
      GridColorFixed  =   14737632
      ForeColorFixed  =   10526880
      ForeColor       =   7368816
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
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
   Begin TradeSkilDemo27.FeaturesPanel FeaturesPanel 
      Height          =   8580
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   15134
   End
End
Attribute VB_Name = "fTradeSkilDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements IStateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================
    
Private Const ModuleName                            As String = "fTradeSkilDemo"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTradeBuildAPI                   As TradeBuildAPI
Attribute mTradeBuildAPI.VB_VarHelpID = -1
Private mConfigStore                                As ConfigurationStore

Private WithEvents mTickers                         As Tickers
Attribute mTickers.VB_VarHelpID = -1

Private mFeaturesPanelHidden                        As Boolean
Private mFeaturesPanelPinned                        As Boolean

Private mInfoPanelHidden                            As Boolean
Private mInfoPanelPinned                            As Boolean

Private mClockDisplay                               As ClockDisplay

Private mAppInstanceConfig                          As ConfigurationSection

Private WithEvents mOrderRecoveryFutureWaiter       As FutureWaiter
Attribute mOrderRecoveryFutureWaiter.VB_VarHelpID = -1
Private WithEvents mChartsCreationFutureWaiter      As FutureWaiter
Attribute mChartsCreationFutureWaiter.VB_VarHelpID = -1

Private WithEvents mContractSelectionHelper         As ContractSelectionHelper
Attribute mContractSelectionHelper.VB_VarHelpID = -1

Private mChartForms                                 As New ChartForms

Private mPreviousMainForm                           As fTradeSkilDemo

Private mOrderTicket                                As fOrderTicket

Private WithEvents mFeaturesPanelForm               As fFeaturesPanel
Attribute mFeaturesPanelForm.VB_VarHelpID = -1
Private WithEvents mInfoPanelForm                   As fInfoPanel
Attribute mInfoPanelForm.VB_VarHelpID = -1

Private mTheme                                      As ITheme

Private mFinishing                                  As Boolean

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitialiseCommonControls
Set mOrderRecoveryFutureWaiter = New FutureWaiter
Set mChartsCreationFutureWaiter = New FutureWaiter
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

LogMessage "Executing Form_Load"

LogMessage "Setting up clock"
Set mClockDisplay = New ClockDisplay
mClockDisplay.Initialise StatusBar1.Panels("datetime"), StatusBar1.Panels("timezone")
mClockDisplay.SetClock getDefaultClock

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_QueryUnload( _
                Cancel As Integer, _
                UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

If UnloadMode <> vbFormCode Then mFinishing = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

Static prevHeight As Long
Static prevWidth As Long

If Me.WindowState = FormWindowStateConstants.vbMinimized Then Exit Sub

If Me.Width < FeaturesPanel.Width + 120 Then Me.Width = FeaturesPanel.Width + 120

StatusBar1.Top = Me.ScaleHeight - StatusBar1.Height

If StatusBar1.Top - 120 < FeaturesPanel.Top + 8700 Then Me.Height = Me.Height + FeaturesPanel.Top + 8700 - StatusBar1.Top + 120

If Me.Width = prevWidth And Me.Height = prevHeight Then Exit Sub

prevWidth = Me.Width
prevHeight = Me.Height

Resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

updateInstanceSettings

LogMessage "Hiding forms"

mChartForms.HideCharts
mChartForms.HideHistoricalCharts

Dim f As Form
For Each f In Forms
    If Not TypeOf f Is fTradeSkilDemo And Not TypeOf f Is fSplash Then f.Hide
Next
Me.Hide

LogMessage "Shutting down clock"
mClockDisplay.Finish

LogMessage "Finishing Features Panel"
FeaturesPanel.Finish
If Not mFeaturesPanelForm Is Nothing Then mFeaturesPanelForm.Finish

LogMessage "Finishing Info Panel"
InfoPanel.Finish
If Not mInfoPanelForm Is Nothing Then mInfoPanelForm.Finish

Shutdown

LogMessage "Closing charts and market depth forms"
closeChartsAndMarketDepthForms

LogMessage "Closing config editor form"
gUnloadConfigEditor

LogMessage "Closing order ticket"
If Not mOrderTicket Is Nothing Then
    Unload mOrderTicket
    Set mOrderTicket = Nothing
End If

LogMessage "Closing other forms"
For Each f In Forms
    If Not TypeOf f Is fTradeSkilDemo And Not TypeOf f Is fSplash Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

LogMessage "Stopping tickers"
If Not mTickers Is Nothing Then mTickers.Finish

LogMessage "Unloading main form"

If mFinishing Then gSetFinished

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' IStateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source

Select Case ev.State
Case MarketDataSourceStates.MarketDataSourceStateCreated

Case MarketDataSourceStates.MarketDataSourceStateReady
    If lDataSource Is getSelectedDataSource Then mClockDisplay.SetClockFuture lDataSource.ClockFuture
Case MarketDataSourceStates.MarketDataSourceStateRunning
    
Case MarketDataSourceStates.MarketDataSourceStatePaused

Case MarketDataSourceStates.MarketDataSourceStateStopped
    If Not getSelectedDataSource Is Nothing Then
        mClockDisplay.SetClockFuture getSelectedDataSource.ClockFuture
    End If
    
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub FeaturesPanel_ConfigsChanged()
Const ProcName As String = "FeaturesPanel_ConfigsChanged"
On Error GoTo Err

mFeaturesPanelForm.SetupCurrentConfigCombo

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_Hide()
Const ProcName As String = "FeaturesPanel_Hide"
On Error GoTo Err

FeaturesPanel.Visible = False
mFeaturesPanelHidden = True
Resize
updateInstanceSettings
ShowFeaturesPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_HistContractSearchCancelled()
Const ProcName As String = "FeaturesPanel_HistContractSearchCancelled"
On Error GoTo Err

mFeaturesPanelForm.CancelHistContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_HistContractSearchCleared()
Const ProcName As String = "FeaturesPanel_HistContractSearchCleared"
On Error GoTo Err

mFeaturesPanelForm.ClearHistContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_HistContractsLoaded(ByVal pContracts As IContracts)
Const ProcName As String = "FeaturesPanel_HistContractsLoaded"
On Error GoTo Err

mFeaturesPanelForm.LoadHistContractsForUserChoice pContracts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_LiveContractSearchCancelled()
Const ProcName As String = "FeaturesPanel_LiveContractSearchCancelled"
On Error GoTo Err

mFeaturesPanelForm.CancelLiveContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_LiveContractSearchCleared()
Const ProcName As String = "FeaturesPanel_LiveContractSearchCleared"
On Error GoTo Err

mFeaturesPanelForm.ClearLiveContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_LiveContractsLoaded(ByVal pContracts As IContracts)
Const ProcName As String = "FeaturesPanel_LiveContractsLoaded"
On Error GoTo Err

mFeaturesPanelForm.LoadLiveContractsForUserChoice pContracts, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FeaturesPanel_Unpin()
Const ProcName As String = "FeaturesPanel_Unpin"
On Error GoTo Err

ShowFeaturesPanelPicture.Visible = True
FeaturesPanel.Visible = False
mFeaturesPanelPinned = False
Resize
updateInstanceSettings

If Not mTheme Is Nothing Then mFeaturesPanelForm.Theme = mTheme
mFeaturesPanelForm.Show vbModeless, Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub InfoPanel_Hide()
Const ProcName As String = "InfoPanel_Hide"
On Error GoTo Err

InfoPanel.Visible = False
mInfoPanelHidden = True
Resize
updateInstanceSettings
ShowInfoPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub InfoPanel_Unpin()
Const ProcName As String = "InfoPanel_Unpin"
On Error GoTo Err

ShowInfoPanelPicture.Visible = True
InfoPanel.Visible = False
mInfoPanelPinned = False
Resize
updateInstanceSettings

If Not mTheme Is Nothing Then mInfoPanelForm.Theme = mTheme
mInfoPanelForm.Show vbModeless, Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ShowFeaturesPanelPicture_Click()
Const ProcName As String = "ShowFeaturesPanelPicture_Click"
On Error GoTo Err

showFeaturesPanel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName

End Sub

Private Sub ShowInfoPanelPicture_Click()
Const ProcName As String = "ShowInfoPanelPicture_Click"
On Error GoTo Err

showInfoPanel

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid1_ErroredTickerRemoved(ByVal pTicker As IMarketDataSource)
Const ProcName As String = "TickerGrid1_ErroredTickerRemoved"
On Error GoTo Err

pTicker.Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TickerGrid1_KeyUp"
On Error GoTo Err

Select Case KeyCode
Case vbKeyDelete
    StopSelectedTickers
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid1_TickerSelectionChanged()
Const ProcName As String = "TickerGrid1_TickerSelectionChanged"
On Error GoTo Err

handleSelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid1_TickerSymbolEntered(ByVal pSymbol As String, ByVal pPreferredRow As Long)
Const ProcName As String = "TickerGrid1_TickerSymbolEntered"
On Error GoTo Err

Set mContractSelectionHelper = CreateContractSelectionHelper( _
                                        CreateContractSpecifierFromString(pSymbol), _
                                        pPreferredRow, _
                                        mTradeBuildAPI.ContractStorePrimary, _
                                        mTradeBuildAPI.ContractStoreSecondary)

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gModelessMsgBox Err.Description, MsgBoxExclamation, mTheme, "Attention"
Else
    gNotifyUnhandledError ProcName, ModuleName
End If
End Sub

'================================================================================
' mChartsCreationFutureWaiter Event Handlers
'================================================================================

Private Sub mChartsCreationFutureWaiter_WaitAllCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mChartsCreationFutureWaiter_WaitAllCompleted"
On Error GoTo Err

loadAppInstanceCompletion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' mContractSelectionHelper Event Handlers
'================================================================================

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
    TickerGrid1.StartTickerFromContract _
                    mContractSelectionHelper.Contracts.ItemAtIndex(1), _
                    mContractSelectionHelper.PreferredTickerGridRow, _
                    mContractSelectionHelper.ContractSpecifier.Expiry
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mContractSelectionHelper_ShowContractSelector()
Const ProcName As String = "mContractSelectionHelper_ShowContractSelector"
On Error GoTo Err

If mFeaturesPanelHidden Then showFeaturesPanel
FeaturesPanel.ShowTickersPane
FeaturesPanel.LoadLiveContractsForUserChoice _
                mContractSelectionHelper.Contracts, _
                mContractSelectionHelper.PreferredTickerGridRow
mFeaturesPanelForm.ShowTickersPane
mFeaturesPanelForm.LoadLiveContractsForUserChoice _
                mContractSelectionHelper.Contracts, _
                mContractSelectionHelper.PreferredTickerGridRow

Exit Sub

Err:
If Err.Number = 401 Then Exit Sub ' Can't show non-modal form when modal form is displayed
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mFeaturesPanelForm Event Handlers
'================================================================================

Private Sub mFeaturesPanelForm_ConfigsChanged()
Const ProcName As String = "mFeaturesPanelForm_ConfigsChanged"
On Error GoTo Err

FeaturesPanel.SetupCurrentConfigCombo

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_Hide()
Const ProcName As String = "mFeaturesPanelForm_Hide"
On Error GoTo Err

mFeaturesPanelForm.Hide
mFeaturesPanelHidden = True
updateInstanceSettings
ShowFeaturesPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_HistContractSearchCancelled()
Const ProcName As String = "mFeaturesPanelForm_HistContractSearchCancelled"
On Error GoTo Err

FeaturesPanel.CancelHistContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_HistContractSearchCleared()
Const ProcName As String = "mFeaturesPanelForm_HistContractSearchCleared"
On Error GoTo Err

FeaturesPanel.ClearHistContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_HistContractsLoaded(ByVal pContracts As IContracts)
Const ProcName As String = "mFeaturesPanelForm_HistContractsLoaded"
On Error GoTo Err

FeaturesPanel.LoadHistContractsForUserChoice pContracts

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_LiveContractSearchCancelled()
Const ProcName As String = "mFeaturesPanelForm_LiveContractSearchCancelled"
On Error GoTo Err

FeaturesPanel.CancelLiveContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_LiveContractSearchCleared()
Const ProcName As String = "mFeaturesPanelForm_LiveContractSearchCleared"
On Error GoTo Err

FeaturesPanel.ClearLiveContractSearch

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_LiveContractsLoaded(ByVal pContracts As IContracts)
Const ProcName As String = "mFeaturesPanelForm_LiveContractsLoaded"
On Error GoTo Err

FeaturesPanel.LoadLiveContractsForUserChoice pContracts, 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mFeaturesPanelForm_Pin()
Const ProcName As String = "mFeaturesPanelForm_Pin"
On Error GoTo Err

mFeaturesPanelForm.Hide
mFeaturesPanelPinned = True
FeaturesPanel.Visible = True
ShowFeaturesPanelPicture.Visible = False
Resize
updateInstanceSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mInfoPanelForm Event Handlers
'================================================================================

Private Sub mInfoPanelForm_Hide()
Const ProcName As String = "mInfoPanelForm_Hide"
On Error GoTo Err

mInfoPanelForm.Hide
mInfoPanelHidden = True
updateInstanceSettings
ShowInfoPanelPicture.Visible = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mInfoPanelForm_Pin()
Const ProcName As String = "mInfoPanelForm_Pin"
On Error GoTo Err

mInfoPanelForm.Hide
mInfoPanelPinned = True
InfoPanel.Visible = True
ShowInfoPanelPicture.Visible = False
Resize
updateInstanceSettings

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mOrderRecoveryFutureWaiter Event Handlers
'================================================================================

Private Sub mOrderRecoveryFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mOrderRecoveryFutureWaiter_WaitCompleted"
On Error GoTo Err

If ev.Future.IsFaulted Then
    LogMessage "Order recovery failed"
ElseIf ev.Future.IsAvailable Then
    
    If Not mPreviousMainForm Is Nothing Then
        Unload mPreviousMainForm
        Set mPreviousMainForm = Nothing
    End If
    
    LogMessage "Order recovery completed    "
    loadAppInstanceConfig
    
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

Dim lTicker As Ticker

Select Case ev.ChangeType
Case CollItemAdded
    Set lTicker = ev.AffectedItem
    lTicker.AddStateChangeListener Me
Case CollItemRemoved
    Set lTicker = ev.AffectedItem
    lTicker.RemoveStateChangeListener Me
Case CollItemChanged

Case CollOrderChanged

Case CollCollectionCleared

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mTradeBuildAPI Event Handlers
'================================================================================

Private Sub mTradeBuildAPI_Notification( _
                ByRef ev As NotificationEventData)
Const ProcName As String = "mTradeBuildAPI_Notification"
On Error GoTo Err

Static sCantConnectNotified As Boolean

Select Case ev.EventCode
Case ApiNotifyCodes.ApiNotifyServiceProviderError
    Dim spError As ServiceProviderError
    Set spError = mTradeBuildAPI.GetServiceProviderError
    LogMessage "Error from " & _
                        spError.ServiceProviderName & _
                        ": code " & spError.ErrorCode & _
                        ": " & spError.Message
Case ApiNotifyCodes.ApiNotifyCantConnect
    If Not sCantConnectNotified Then
        gModelessMsgBox ev.EventMessage, MsgBoxCritical, mTheme, "Can't connect"
        sCantConnectNotified = True
    End If
Case ApiNotifyCodes.ApiNotifyConnected
    sCantConnectNotified = False
Case Else
    LogMessage "Notification: code=" & ev.EventCode & "; source=" & TypeName(ev.Source) & ": " & _
                ev.EventMessage & vbCrLf
End Select

Exit Sub

Err:
If Err.Number = 401 Then Exit Sub ' Can't show non-modal form when modal form is displayed
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub ApplyTheme(ByVal pThemeName As String)
Const ProcName As String = "ApplyTheme"
On Error GoTo Err

Static sThemeName As String

If UCase$(pThemeName) = UCase$(sThemeName) Then Exit Sub
sThemeName = UCase$(pThemeName)

If sThemeName = "BLACK" Then
    Set mTheme = New BlackTheme
ElseIf sThemeName = "BLUE" Then
    Set mTheme = New BlueTheme
ElseIf sThemeName = "NATIVE" Then
    Set mTheme = New NativeTheme
Else
    Set mTheme = New BlackTheme
End If

LogMessage "Applying theme to main form", LogLevelHighDetail

mAppInstanceConfig.SetSetting ConfigSettingCurrentTheme, sThemeName

Me.BackColor = mTheme.BaseColor
gApplyTheme mTheme, Me.Controls

ShowFeaturesPanelPicture.BackColor = mTheme.BaseColor
ShowInfoPanelPicture.BackColor = mTheme.BaseColor

Dim lForm As Object
For Each lForm In Forms
    If TypeOf lForm Is IThemeable Then
        LogMessage "Applying theme to form: " & lForm.caption, LogLevelHighDetail
        lForm.Theme = mTheme
    End If
Next

SendMessage StatusBar1.hWnd, SB_SETBKCOLOR, 0, NormalizeColor(mTheme.StatusBarBackColor)

Dim lhDC As Long
lhDC = GetDC(StatusBar1.hWnd)
SetTextColor lhDC, NormalizeColor(mTheme.StatusBarForeColor)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Initialise( _
                ByVal pTradeBuildAPI As TradeBuildAPI, _
                ByVal pConfigStore As ConfigurationStore, _
                ByVal pAppInstanceConfig As ConfigurationSection, _
                ByRef pErrorMessage As String) As Boolean
Const ProcName As String = "initialise"
On Error GoTo Err

Set mTradeBuildAPI = pTradeBuildAPI
Set mConfigStore = pConfigStore
Set mAppInstanceConfig = pAppInstanceConfig
Set mPreviousMainForm = gMainForm

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionApplication

Set mTickers = mTradeBuildAPI.Tickers
If mTickers Is Nothing Then
    pErrorMessage = "No tickers object is available: one or more service providers may be missing or disabled"
    Initialise = False
    Exit Function
End If

LogMessage "Recovering orders from last session"
mOrderRecoveryFutureWaiter.Add CreateFutureFromTask(mTradeBuildAPI.RecoverOrders())

Initialise = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Shutdown()
Const ProcName As String = "Shutdown"
On Error GoTo Err

Static sAlreadyShutdown As Boolean
If sAlreadyShutdown Then Exit Sub

sAlreadyShutdown = True

LogMessage "Finishing UI controls"
finishUIControls

LogMessage "Removing service providers"
mTradeBuildAPI.ServiceProviders.RemoveAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub applyInstanceSettings()
Const ProcName As String = "applyInstanceSettings"
On Error GoTo Err

LogMessage "Loading configuration: positioning main form"
Select Case mAppInstanceConfig.GetSetting(ConfigSettingMainFormWindowstate, WindowStateNormal)
Case WindowStateMaximized
    Me.WindowState = FormWindowStateConstants.vbMaximized
Case WindowStateMinimized
    Me.WindowState = FormWindowStateConstants.vbMinimized
Case WindowStateNormal
    Me.Left = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormLeft, 0)) * Screen.TwipsPerPixelX
    Me.Top = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormTop, 0)) * Screen.TwipsPerPixelY
    Me.Width = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX)) * Screen.TwipsPerPixelX
    Me.Height = CLng(mAppInstanceConfig.GetSetting(ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelectedTickers()
Const ProcName As String = "clearSelectedTickers"
On Error GoTo Err

TickerGrid1.DeselectSelectedTickers
handleSelectedTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub closeChartsAndMarketDepthForms()
Const ProcName As String = "closeChartsAndMarketDepthForms"
On Error GoTo Err

mChartForms.Finish

Dim f As Form
For Each f In Forms
    If TypeOf f Is fMarketDepth Then
        LogMessage "Closing form: caption=" & f.caption & "; type=" & TypeName(f)
        Unload f
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub finishUIControls()
Const ProcName As String = "finishUIControls"
On Error GoTo Err

TickerGrid1.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDefaultClock() As Clock
Const ProcName As String = "getDefaultClock"
On Error GoTo Err

Static sClock As Clock
If sClock Is Nothing Then Set sClock = GetClock("") ' create a clock running local time
Set getDefaultClock = sClock

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelectedDataSource() As IMarketDataSource
Const ProcName As String = "getSelectedDataSource"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 1 Then Set getSelectedDataSource = TickerGrid1.SelectedTickers.Item(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub handleSelectedTickers()
Const ProcName As String = "handleSelectedTickers"
On Error GoTo Err

If TickerGrid1.SelectedTickers.Count = 0 Then
    mClockDisplay.SetClock getDefaultClock
Else
    Dim lTicker As Ticker
    Set lTicker = getSelectedDataSource
    If lTicker Is Nothing Then
        mClockDisplay.SetClock getDefaultClock
    ElseIf lTicker.State = MarketDataSourceStateRunning Then
        mClockDisplay.SetClockFuture lTicker.ClockFuture
    Else
        mClockDisplay.SetClock getDefaultClock
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadAppInstanceConfig()
Const ProcName As String = "loadAppInstanceConfig"
On Error GoTo Err

LogMessage "Setting application title"
Me.caption = gAppTitle & _
            " - " & mAppInstanceConfig.InstanceQualifier

LogMessage "Loading configuration: " & mAppInstanceConfig.InstanceQualifier

LogMessage "Loading configuration: Setting up ticker grid"
setupTickerGrid

LogMessage "Loading configuration: setting up order ticket"
setupOrderTicket

applyInstanceSettings

LogMessage "Loading configuration: loading tickers into ticker grid"
TickerGrid1.LoadFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionTickerGrid)

LogMessage "Loading configuration: loading default study configurations"
LoadDefaultStudyConfigurationsFromConfig mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionDefaultStudyConfigs)

LogMessage "Loading configuration: creating charts"
mChartsCreationFutureWaiter.Add startCharts

LogMessage "Loading configuration: creating historical charts"
mChartsCreationFutureWaiter.Add startHistoricalCharts

LogMessage "Loading configuration: initialising Info Panels"
setupInfoPanels

LogMessage "Loading configuration: initialising Features Panels"
setupFeaturesPanels

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadAppInstanceCompletion()
Const ProcName As String = "loadAppInstanceCompletion"
On Error GoTo Err

LogMessage "Loading configuration: applying theme"
ApplyTheme mAppInstanceConfig.GetSetting(ConfigSettingCurrentTheme, "Black")

LogMessage "Loading configuration: showing main form"
Me.Show vbModeless

LogMessage "Loading configuration: showing charts"
mChartForms.ShowCharts gMainForm

LogMessage "Loading configuration: showing historical charts"
mChartForms.ShowHistoricalCharts gMainForm

LogMessage "Loading configuration: applying theme to charts"
mChartForms.Theme = mTheme

LogMessage "Loading configuration: showing Features and Info panels"
If Not mFeaturesPanelHidden Then showFeaturesPanel
If Not mInfoPanelHidden Then showInfoPanel

LogMessage "Loaded configuration: " & mAppInstanceConfig.InstanceQualifier

Me.SetFocus
gSplashScreen.Hide

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Resize()
Const ProcName As String = "Resize"
On Error GoTo Err

Dim lLeft As Long
If mFeaturesPanelHidden Or Not mFeaturesPanelPinned Then
    lLeft = ShowFeaturesPanelPicture.Width + 60
Else
    lLeft = 120 + FeaturesPanel.Width + 120
End If

FeaturesPanel.Height = StatusBar1.Top - FeaturesPanel.Top - 120

If Not mInfoPanelHidden And mInfoPanelPinned Then
    InfoPanel.Move lLeft, _
                        StatusBar1.Top - InfoPanel.Height - 120, _
                        Me.ScaleWidth - lLeft - 120
End If

ShowInfoPanelPicture.Move Me.ScaleWidth - 345, _
                        StatusBar1.Top - 240

TickerGrid1.Move lLeft, _
                TickerGrid1.Top, _
                Me.ScaleWidth - lLeft - 120, _
                IIf(Not mInfoPanelHidden And mInfoPanelPinned, InfoPanel.Top - 120, ShowInfoPanelPicture.Top - 60) - TickerGrid1.Top

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupFeaturesPanels()
Const ProcName As String = "setupFeaturesPanels"
On Error GoTo Err

LogMessage "Initialising fixed Features Panel"
FeaturesPanel.Initialise True, mTradeBuildAPI, mConfigStore, mAppInstanceConfig, TickerGrid1, InfoPanel, mInfoPanelForm.InfoPanel, mChartForms, mOrderTicket
mFeaturesPanelPinned = CBool(mAppInstanceConfig.GetSetting(ConfigSettingFeaturesPanelPinned, "True"))
mFeaturesPanelHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingFeaturesPanelHidden, "False"))
    
LogMessage "Creating floating Features Panel"
Set mFeaturesPanelForm = New fFeaturesPanel
LogMessage "Initialising floating Features Panel"
mFeaturesPanelForm.Initialise mTradeBuildAPI, mConfigStore, mAppInstanceConfig, TickerGrid1, InfoPanel, mInfoPanelForm.InfoPanel, mChartForms, mOrderTicket

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupInfoPanels()
Const ProcName As String = "setupInfoPanels"
On Error GoTo Err

InfoPanel.Initialise True, mTradeBuildAPI, mAppInstanceConfig, TickerGrid1, mOrderTicket
mInfoPanelPinned = CBool(mAppInstanceConfig.GetSetting(ConfigSettingInfoPanelPinned, "True"))
mInfoPanelHidden = CBool(mAppInstanceConfig.GetSetting(ConfigSettingInfoPanelHidden, "False"))
    
Set mInfoPanelForm = New fInfoPanel
mInfoPanelForm.Initialise mTradeBuildAPI, mAppInstanceConfig, TickerGrid1, mOrderTicket

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupOrderTicket()
Const ProcName As String = "setupOrderTicket"
On Error GoTo Err

Set mOrderTicket = New fOrderTicket
mOrderTicket.Initialise mAppInstanceConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickerGrid()
Const ProcName As String = "setupTickerGrid"
On Error GoTo Err

TickerGrid1.Initialise mTickers

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showFeaturesPanel()
Const ProcName As String = "showFeaturesPanel"
On Error GoTo Err

mFeaturesPanelHidden = False
If mFeaturesPanelPinned Then
    ShowFeaturesPanelPicture.Visible = False
    FeaturesPanel.Visible = True
    Resize
    Me.Refresh
Else
    Static sDoneFirstShow As Boolean
    If Not sDoneFirstShow Then
        sDoneFirstShow = True
        mFeaturesPanelForm.Move CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelLeft, 0)) * Screen.TwipsPerPixelX, _
                CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelTop, (Screen.Height - Me.Height) / Screen.TwipsPerPixelY)) * Screen.TwipsPerPixelY, _
                CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelWidth, 280)) * Screen.TwipsPerPixelX, _
                CLng(mAppInstanceConfig.GetSetting(ConfigSettingFloatingFeaturesPanelHeight, 650)) * Screen.TwipsPerPixelY
    End If
    ShowFeaturesPanelPicture.Visible = True
    mFeaturesPanelForm.Show vbModeless, Me
    If Not mTheme Is Nothing Then mFeaturesPanelForm.Theme = mTheme
End If
updateInstanceSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showInfoPanel()
Const ProcName As String = "showInfoPanel"
On Error GoTo Err

mInfoPanelHidden = False
If mInfoPanelPinned Then
    ShowInfoPanelPicture.Visible = False
    InfoPanel.Visible = True
    Resize
    Me.Refresh
Else
    ShowInfoPanelPicture.Visible = True
    mInfoPanelForm.Show vbModeless, Me
    If Not mTheme Is Nothing Then mInfoPanelForm.Theme = mTheme
End If
updateInstanceSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function startCharts() As IFuture
Const ProcName As String = "startCharts"
On Error GoTo Err

Set startCharts = CreateFutureFromTask( _
                        mChartForms.LoadChartsFromConfigAsync( _
                                mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionCharts), _
                                mTickers, _
                                mTradeBuildAPI.BarFormatterLibManager, _
                                mTradeBuildAPI.HistoricalDataStoreInput.TimePeriodValidator))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function startHistoricalCharts() As IFuture
Const ProcName As String = "startHistoricalCharts"
On Error GoTo Err

Set startHistoricalCharts = CreateFutureFromTask( _
                        mChartForms.LoadHistoricalChartsFromConfigAsync( _
                                mAppInstanceConfig.AddPrivateConfigurationSection(ConfigSectionHistoricCharts), _
                                mTradeBuildAPI.StudyLibraryManager, _
                                mTradeBuildAPI.HistoricalDataStoreInput, _
                                mTradeBuildAPI.BarFormatterLibManager))
                    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTickers As SelectedTickers
Set lTickers = TickerGrid1.SelectedTickers

TickerGrid1.StopSelectedTickers

Dim lTicker As IMarketDataSource
For Each lTicker In lTickers
    lTicker.Finish
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub updateInstanceSettings()
Const ProcName As String = "updateInstanceSettings"
On Error GoTo Err

If mAppInstanceConfig Is Nothing Then Exit Sub

mAppInstanceConfig.AddPrivateConfigurationSection ConfigSectionMainForm
Select Case Me.WindowState
Case FormWindowStateConstants.vbMaximized
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMaximized
Case FormWindowStateConstants.vbMinimized
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateMinimized
Case FormWindowStateConstants.vbNormal
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWindowstate, WindowStateNormal
    mAppInstanceConfig.SetSetting ConfigSettingMainFormLeft, Me.Left / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingMainFormTop, Me.Top / Screen.TwipsPerPixelY
    mAppInstanceConfig.SetSetting ConfigSettingMainFormWidth, Me.Width / Screen.TwipsPerPixelX
    mAppInstanceConfig.SetSetting ConfigSettingMainFormHeight, Me.Height / Screen.TwipsPerPixelY
End Select

mAppInstanceConfig.SetSetting ConfigSettingFeaturesPanelHidden, CStr(mFeaturesPanelHidden)
mAppInstanceConfig.SetSetting ConfigSettingFeaturesPanelPinned, CStr(mFeaturesPanelPinned)

mAppInstanceConfig.SetSetting ConfigSettingInfoPanelHidden, CStr(mInfoPanelHidden)
mAppInstanceConfig.SetSetting ConfigSettingInfoPanelPinned, CStr(mInfoPanelPinned)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


