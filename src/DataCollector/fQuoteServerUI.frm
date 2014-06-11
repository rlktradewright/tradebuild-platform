VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form fDataCollectorUI 
   Caption         =   "TradeBuild Data Collector Version 2.7"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   23
      Top             =   3855
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8784
         EndProperty
      EndProperty
   End
   Begin VB.TextBox SecsSinceLastTickText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      ToolTipText     =   "The number of seconds since the last tick received"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox TicksPerSecText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "The number of ticks received during the last second"
      Top             =   360
      Width           =   615
   End
   Begin VB.CommandButton ShowHideMonitorButton 
      Caption         =   "Hide activity monitor"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      ToolTipText     =   "Shows or hides the activity monitor"
      Top             =   360
      Width           =   1695
   End
   Begin TabDlg.SSTab ActivityMonitor 
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   5530
      _Version        =   393216
      Style           =   1
      TabHeight       =   494
      ForeColor       =   -2147483630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Activity"
      TabPicture(0)   =   "fQuoteServerUI.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TickerScroll"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TickersContainerPicture"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "&Log"
      TabPicture(1)   =   "fQuoteServerUI.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "LogText"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Configuration"
      TabPicture(2)   =   "fQuoteServerUI.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "ConfigDetailsButton"
      Tab(2).Control(1)=   "ConfigNameText"
      Tab(2).Control(2)=   "Label4"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton ConfigDetailsButton 
         Caption         =   "Details..."
         Height          =   375
         Left            =   -71640
         TabIndex        =   26
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox ConfigNameText 
         Height          =   285
         Left            =   -74400
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   960
         Width           =   3495
      End
      Begin VB.PictureBox TickersContainerPicture 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2625
         ScaleWidth      =   4785
         TabIndex        =   7
         Top             =   360
         Width           =   4815
         Begin VB.PictureBox TickersPicture 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            ScaleHeight     =   285
            ScaleWidth      =   4815
            TabIndex        =   8
            Top             =   0
            Width           =   4815
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   4
               Left            =   3840
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   0
               Left            =   0
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   12
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   1
               Left            =   960
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   11
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   2
               Left            =   1920
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   10
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.TextBox ShortNameText 
               Height          =   285
               Index           =   3
               Left            =   2880
               Locked          =   -1  'True
               MousePointer    =   1  'Arrow
               TabIndex        =   9
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   4
               Left            =   4560
               TabIndex        =   18
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   0
               Left            =   720
               TabIndex        =   17
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   3
               Left            =   3600
               TabIndex        =   16
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   2
               Left            =   2640
               TabIndex        =   15
               Top             =   0
               Width           =   255
            End
            Begin VB.Label DataLightLabel 
               Height          =   285
               Index           =   1
               Left            =   1680
               TabIndex        =   14
               Top             =   0
               Width           =   255
            End
         End
      End
      Begin VB.TextBox LogText 
         Height          =   2655
         Left            =   -74880
         MaxLength       =   65535
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   360
         Width           =   5055
      End
      Begin VB.VScrollBar TickerScroll 
         Height          =   2700
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "Current configuration:"
         Height          =   375
         Left            =   -74760
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.CommandButton StartStopButton 
      Caption         =   "Start"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      ToolTipText     =   "Starts or stops data collection"
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox ConnectionStatusText 
      BackColor       =   &H8000000F&
      Height          =   255
      Left            =   3480
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Indicates the health of the connection to the realtime data source: green is ok, red is error"
      Top             =   120
      Width           =   255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Secs no data"
      Height          =   255
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ticks per sec"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection status"
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "fDataCollectorUI"
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
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements DeferredAction
Implements IBarOutputMonitor
Implements IGenericTickListener
Implements IRawMarketDepthListener
Implements StateChangeListener
Implements ITickfileOutputMonitor
Implements LogListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "fDataCollectorUI"

Private Const TickerScrollMax As Integer = 32767
Private Const TickerScrollMin As Integer = 0

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type TickerTableEntry
    Ticker                  As Ticker
    tli                     As TimerListItem
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mDataCollector As DataCollector
Attribute mDataCollector.VB_VarHelpID = -1

Private mTickers() As TickerTableEntry

Private mTimerList As TimerList

Private mLastTickTime As Date
Private mNoDataRestartSecs As Long

Private WithEvents mClock As Clock
Attribute mClock.VB_VarHelpID = -1

Private mNumTicksSinceConnected As Long
Private mNumTicksThisSecond As Long

Private mActivityMonitorVisible As Boolean

Private mCollectingData As Boolean

Private mAdjustingSize As Boolean
Private mCurrentHeight As Long
Private mCurrentWidth As Long

Private mStartStopButtonInitialLeft As Long

Private mLinesToScroll As Integer

Private mLineSpacing As Integer

Private mStartStopTimePanel As Panel

Private mFormatter As LogFormatter

Private mConfigManager As ConfigManager

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mConnected As Boolean

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

InitCommonControls
Set mTimerList = GetGlobalTimerList
ReDim mTickers(99) As TickerTableEntry
Set mFormatter = CreateBasicLogFormatter(TimestampTimeOnlyLocal)
GetLogger("log").AddLogListener Me
Set mFutureWaiter = New FutureWaiter

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
On Error GoTo Err

mStartStopButtonInitialLeft = StartStopButton.Left
TickerScroll.Min = TickerScrollMin
TickerScroll.Max = TickerScrollMax
mLineSpacing = ShortNameText(0).Height - Screen.TwipsPerPixelY
mCurrentHeight = Me.Height
mCurrentWidth = Me.Width

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Const ProcName As String = "Form_QueryUnload"
On Error GoTo Err

Select Case UnloadMode
Case QueryUnloadConstants.vbAppTaskManager
Case QueryUnloadConstants.vbAppWindows
Case QueryUnloadConstants.vbFormCode
Case QueryUnloadConstants.vbFormControlMenu
    If mCollectingData Then
        If MsgBox("Please confirm that you wish to stop data collection", _
                    vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then Cancel = True
    End If
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

If Me.WindowState = vbMinimized Then Exit Sub
If Me.Height <> mCurrentHeight Then resizeHeight
If Me.Width <> mCurrentWidth Then resizeWidth

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Terminate()
Const ProcName As String = "Form_Terminate"
On Error GoTo Err

TerminateTWUtilities

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

LogMessage "Data Collector program exiting"
GetLogger("log").RemoveLogListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "DeferredAction_Run"
On Error GoTo Err

startCollecting "Restarting collection"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IBarOutputMonitor Interface Members
'================================================================================

Private Sub IBarOutputMonitor_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "IBarOutputMonitor_NotifyEvent"
On Error GoTo Err

Dim lWriter As IBarWriter
Set lWriter = ev.Source
LogMessage "Bar writer notification (" & ev.EventCode & "): " & ev.EventMessage & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyNotReady(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyNotReady"
On Error GoTo Err

Dim lWriter As IBarWriter
Set lWriter = pSource
LogMessage "Bar writer not ready" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyOutputFileClosed(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyOutputFileClosed"
On Error GoTo Err

Dim lWriter As IBarWriter
Set lWriter = pSource
LogMessage "Bar writer closed file" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyOutputFileCreated(ByVal pSource As Object, ByVal pFilename As String)
Const ProcName As String = "IBarOutputMonitor_NotifyOutputFileCreated"
On Error GoTo Err

Dim lWriter As IBarWriter
Set lWriter = pSource
LogMessage "Writing bars to: " & pFilename & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IBarOutputMonitor_NotifyReady(ByVal pSource As Object)
Const ProcName As String = "IBarOutputMonitor_NotifyReady"
On Error GoTo Err

Dim lWriter As IBarWriter
Set lWriter = pSource
LogMessage "Bar writer ready" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IGenericTickListener Interface Members
'================================================================================

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source
processTickEvent lDataSource.Handle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' IRawMarketDepth Interface Members
'================================================================================

Private Sub IRawMarketDepthListener_MarketDepthNotAvailable(ByVal reason As String)
Const ProcName As String = "IRawMarketDepthListener_MarketDepthNotAvailable"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRawMarketDepthListener_resetMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "IRawMarketDepthListener_resetMarketDepth"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IRawMarketDepthListener_updateMarketDepth(ev As RawMarketDepthEventData)
Const ProcName As String = "IRawMarketDepthListener_updateMarketDepth"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source
processTickEvent lDataSource.Handle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' ITickfileOutputMonitor Interface Members
'================================================================================

Private Sub ITickfileOutputMonitor_NotifyEvent(ev As NotificationEventData)
Const ProcName As String = "ITickfileOutputMonitor_NotifyEvent"
On Error GoTo Err

Dim lWriter As ITickfileWriter
Set lWriter = ev.Source
LogMessage "Tick writer notification (" & ev.EventCode & "): " & ev.EventMessage & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyNotReady(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyNotReady"
On Error GoTo Err

Dim lWriter As ITickfileWriter
Set lWriter = pSource
LogMessage "Tick writer not ready" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileClosed(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyOutputFileClosed"
On Error GoTo Err

Dim lWriter As ITickfileWriter
Set lWriter = pSource
LogMessage "Tick writer closed file" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyOutputFileCreated(ByVal pSource As Object, ByVal pFilename As String)
Const ProcName As String = "ITickfileOutputMonitor_NotifyOutputFileCreated"
On Error GoTo Err

Dim lWriter As ITickfileWriter
Set lWriter = pSource
LogMessage "Writing ticks to: " & pFilename & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITickfileOutputMonitor_NotifyReady(ByVal pSource As Object)
Const ProcName As String = "ITickfileOutputMonitor_NotifyReady"
On Error GoTo Err

Dim lWriter As ITickfileWriter
Set lWriter = pSource
LogMessage "Tick writer ready" & getContractString(lWriter.ContractFuture)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_finish()
'nothing to do
Const ProcName As String = "LogListener_finish"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LogListener_Notify(ByVal logrec As LogRecord)
Const ProcName As String = "LogListener_Notify"
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
LogText.SelText = mFormatter.FormatRecord(logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim tli As TimerListItem
Set tli = ev.Source
If Not tli Is Nothing Then
    If ev.State = TimerListItemStates.TimerListItemStateExpired Then
        switchDataLightOff tli.Data
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub ConfigDetailsButton_Click()
Const ProcName As String = "ConfigDetailsButton_Click"
On Error GoTo Err

Dim f As New fConfig
f.Initialise mConfigManager, True
f.Show vbModeless

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ShowHideMonitorButton_Click()
Const ProcName As String = "ShowHideMonitorButton_Click"
On Error GoTo Err

If mActivityMonitorVisible Then
    hideActivityMonitor
Else
    showActivityMonitor
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub StartStopButton_Click()
Const ProcName As String = "StartStopButton_Click"
On Error GoTo Err

If mCollectingData Then
    stopCollecting "Data collection stopped by user", True
Else
    startCollecting "Data collection started by user"
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerScroll_Change()
Const ProcName As String = "TickerScroll_Change"
On Error GoTo Err

scrollTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mClock Event Handlers
'================================================================================

Private Sub mClock_Tick()
Const ProcName As String = "mClock_Tick"
On Error GoTo Err

TicksPerSecText = mNumTicksThisSecond
TicksPerSecText.Refresh
mNumTicksThisSecond = 0

Dim lSecsSinceLastTick As Long
lSecsSinceLastTick = Int(86400 * (GetTimestamp - mLastTickTime))
SecsSinceLastTickText = lSecsSinceLastTick
SecsSinceLastTickText.Refresh

If mNoDataRestartSecs > 0 And mConnected And lSecsSinceLastTick >= mNoDataRestartSecs And mNumTicksSinceConnected > 0 Then
    Set mClock = Nothing
    stopCollecting "Stopping collection: possible undetected loss of connection to provider", False
    
    LogMessage "Restarting collection in 10 seconds"
    DeferAction Me, , 10, ExpiryTimeUnitSeconds
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mDataCollector Event Handlers
'================================================================================

Private Sub mDataCollector_CollectionStarted()
Const ProcName As String = "mDataCollector_CollectionStarted"
On Error GoTo Err

Dim s As String
If mDataCollector.nextEndTime <> 0 Then
    s = "Collection end: " & _
        FormatDateTime(mDataCollector.nextEndTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.nextEndTime, vbShortTime)
End If
If mDataCollector.exitProgramTime <> 0 Then
    s = IIf(s = "", "P", s & "; p") & _
        "rogram exit: " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
End If

mStartStopTimePanel.Text = s

setStarted

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_CollectionStopped()
Const ProcName As String = "mDataCollector_CollectionStopped"
On Error GoTo Err

Dim s As String
If mDataCollector.nextStartTime <> 0 Then
    s = "Collection start: " & _
        FormatDateTime(mDataCollector.nextStartTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.nextStartTime, vbShortTime)
End If
If mDataCollector.exitProgramTime <> 0 Then
    s = IIf(s = "", "P", s & "; p") & _
        "rogram exit: " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
        FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
End If

mStartStopTimePanel.Text = s

setStopped

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_connected()
Const ProcName As String = "mDataCollector_connected"
On Error GoTo Err

mNumTicksSinceConnected = 0

mConnected = True
ConnectionStatusText.BackColor = vbGreen
StartStopButton.enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_connectFailed(ByVal description As String)
Const ProcName As String = "mDataCollector_connectFailed"
On Error GoTo Err

mConnected = False
ConnectionStatusText.BackColor = vbRed

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_ConnectionClosed()
Const ProcName As String = "mDataCollector_ConnectionClosed"
On Error GoTo Err

mConnected = False
ConnectionStatusText.BackColor = vbRed
StartStopButton.enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_Error(ev As ErrorEventData)
Const ProcName As String = "mDataCollector_Error"
On Error GoTo Err

LogMessage "Error " & ev.ErrorCode & ": " & vbCrLf & _
            ev.ErrorMessage, _
            LogLevelSevere

mConnected = False
stopCollecting "Closing due to error", False
Unload Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_ExitProgram()
Const ProcName As String = "mDataCollector_ExitProgram"
On Error GoTo Err

Unload Me

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_FatalError(ev As ErrorEventData)
Const ProcName As String = "mDataCollector_FatalError"
On Error GoTo Err

Err.Raise ev.ErrorCode, ev.ErrorSource, ev.ErrorMessage

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_Reconnecting()
Const ProcName As String = "mDataCollector_Reconnecting"
On Error GoTo Err

mConnected = False
ConnectionStatusText.BackColor = vbRed
StartStopButton.enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mDataCollector_TickerAdded(ByVal pTicker As Ticker)
Const ProcName As String = "mDataCollector_TickerAdded"
On Error GoTo Err

Dim index As Long
index = pTicker.Handle

If index > UBound(mTickers) Then
    ReDim Preserve mTickers(index / 100 * 100 + 99) As TickerTableEntry
End If
Set mTickers(index).Ticker = pTicker
Set mTickers(index).tli = Nothing

If index > ShortNameText.ubound Then
    Dim i As Long
    For i = ShortNameText.ubound + 1 To index
        Load ShortNameText(i)
        ShortNameText(i).Left = ShortNameText(i - 5).Left
        ShortNameText(i).Top = ShortNameText(i - 5).Top + mLineSpacing
        ShortNameText(i).ZOrder 0
        If i Mod 5 = 0 Then TickersPicture.Height = ShortNameText(i).Top + ShortNameText(i).Height
        
        Load DataLightLabel(i)
        DataLightLabel(i).Left = DataLightLabel(i - 5).Left
        DataLightLabel(i).Top = DataLightLabel(i - 5).Top + mLineSpacing
        DataLightLabel(i).ZOrder 0
        
        setupTickerScroll
    Next
End If

mFutureWaiter.Add pTicker.ContractFuture, index
ShortNameText(index).Visible = True
DataLightLabel(index).Visible = True

pTicker.AddGenericTickListener Me
pTicker.AddRawMarketDepthListener Me

Me.Refresh

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'================================================================================
' mFutureWaiter `Event Handlers
'================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

If Not ev.Future.IsAvailable Then Exit Sub

Dim lContract As IContract
Set lContract = ev.Future.Value
Dim lIndex As Long
lIndex = ev.ContinuationData

ShortNameText(lIndex) = lContract.Specifier.localSymbol
ShortNameText(lIndex).ToolTipText = lContract.Specifier.localSymbol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub Initialise( _
                ByVal pDataCollector As DataCollector, _
                ByVal pconfigManager As ConfigManager, _
                ByVal configName As String, _
                ByVal noAutoStart As Boolean, _
                ByVal showMonitor As Boolean, _
                ByVal pNoDataRestartSecs As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mNoDataRestartSecs = pNoDataRestartSecs
Set mStartStopTimePanel = StatusBar1.Panels.Item(1)

Set mDataCollector = pDataCollector

Set mConfigManager = pconfigManager
ConfigNameText = configName

If showMonitor Then
    mActivityMonitorVisible = True
Else
    hideActivityMonitor
End If

If noAutoStart Then
    If mDataCollector.exitProgramTime <> 0 Then
        mStartStopTimePanel.Text = "Program exit: " & _
                                    FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
                                    FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
    End If
Else
Dim s As String
    If mDataCollector.nextStartTime <> 0 Then
        s = "Collection start: " & _
            FormatDateTime(mDataCollector.nextStartTime, vbShortDate) & " " & _
            FormatDateTime(mDataCollector.nextStartTime, vbShortTime)
    End If
    If mDataCollector.exitProgramTime <> 0 Then
        s = IIf(s = "", "P", s & "; p") & _
            "rogram exit: " & _
            FormatDateTime(mDataCollector.exitProgramTime, vbShortDate) & " " & _
            FormatDateTime(mDataCollector.exitProgramTime, vbShortTime)
    End If
    mStartStopTimePanel.Text = s
    If mDataCollector.nextStartTime = 0 Then
        startCollecting "Data collection started automatically"
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearTickers()
Const ProcName As String = "clearTickers"
On Error GoTo Err

Dim i As Long
For i = 0 To ShortNameText.ubound
    ShortNameText(i).Text = ""
    DataLightLabel(i).BackColor = vbButtonFace
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generateTaskInfo( _
                ByVal en As Enumerator) As String
Const ProcName As String = "generateTaskInfo"
On Error GoTo Err

Dim s As String

Do While en.MoveNext
    Dim ts As TaskSummary
    ts = en.Current
    s = s & "Name: " & ts.Name & _
        "; Priority: " & ts.Priority & _
        "; Start time: " & FormatTimestamp(ts.StartTime, TimestampDateAndTimeISO8601) & _
        "; Last run time: " & FormatTimestamp(ts.LastRunTime, TimestampDateAndTimeISO8601) & _
        "; CPU time: " & Format(ts.TotalCPUTime, "0.0") & vbCrLf
        
Loop
    
generateTaskInfo = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getContractString(ByVal pContractFuture As IFuture) As String
Const ProcName As String = "getContractString"
On Error GoTo Err

If pContractFuture.IsAvailable Then
    Dim lContract As IContract
    Set lContract = pContractFuture.Value
    getContractString = " (" & lContract.Specifier.ToString & ")"
Else
    getContractString = ""
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideActivityMonitor()
Const ProcName As String = "hideActivityMonitor"
On Error GoTo Err

ShowHideMonitorButton.Caption = "Show activity monitor"
mAdjustingSize = True
Me.Height = Me.Height - ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = False
mActivityMonitorVisible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
    
Private Sub processTickEvent( _
                pHandle As Long)
Const ProcName As String = "processTickEvent"
On Error GoTo Err

switchDataLightOn pHandle
mLastTickTime = GetTimestamp
mNumTicksThisSecond = mNumTicksThisSecond + 1
mNumTicksSinceConnected = mNumTicksSinceConnected + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resizeHeight()
Const ProcName As String = "resizeHeight"
On Error GoTo Err

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Height = mCurrentHeight
    Exit Sub
End If

Dim heightIncrement As Long
heightIncrement = Me.Height - mCurrentHeight

If Not mAdjustingSize Then
    If TickersContainerPicture.Height + heightIncrement <= 0 Then
        Me.Height = mCurrentHeight
        Exit Sub
    End If
    
    ActivityMonitor.Height = ActivityMonitor.Height + heightIncrement

    TickersContainerPicture.Height = TickersContainerPicture.Height + heightIncrement

    TickerScroll.Height = TickerScroll.Height + heightIncrement

    LogText.Height = LogText.Height + heightIncrement
    
End If

mCurrentHeight = Me.Height
setupTickerScroll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resizeWidth()
Const ProcName As String = "resizeWidth"
On Error GoTo Err

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

If Me.Width <= mStartStopButtonInitialLeft + StartStopButton.Width + 120 Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

Dim widthIncrement As Long
widthIncrement = Me.Width - mCurrentWidth

If Not mAdjustingSize Then
    
    ActivityMonitor.Width = ActivityMonitor.Width + widthIncrement

    TickersPicture.Width = TickersPicture.Width + widthIncrement

    TickerScroll.Left = TickerScroll.Left + widthIncrement

    LogText.Width = LogText.Width + widthIncrement

    StartStopButton.Left = StartStopButton.Left + widthIncrement
End If

mCurrentWidth = Me.Width

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub scrollTickers()
Const ProcName As String = "scrollTickers"
On Error GoTo Err

If TickersPicture.Height <= TickersContainerPicture.Height Then
    TickersPicture.Top = 0
ElseIf TickerScroll.Value = TickerScrollMax Then
    TickersPicture.Top = -mLinesToScroll * mLineSpacing
Else
    TickersPicture.Top = -Round((mLinesToScroll / TickerScrollMax) * TickerScroll.Value, 0) * mLineSpacing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setStarted()
Const ProcName As String = "setStarted"
On Error GoTo Err

mCollectingData = True
StartStopButton.Caption = "Stop"
StartStopButton.enabled = True

mLastTickTime = GetTimestamp

Set mClock = GetClock

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setStopped()
Const ProcName As String = "setStopped"
On Error GoTo Err

Set mClock = Nothing

mCollectingData = False
StartStopButton.Caption = "Start"

mConnected = False
ConnectionStatusText.BackColor = vbButtonFace

clearTickers

Set mClock = Nothing

SecsSinceLastTickText = ""
TicksPerSecText = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickerScroll()
Const ProcName As String = "setupTickerScroll"
On Error GoTo Err

Dim totalLines As Long
totalLines = (ShortNameText.ubound + 5) / 5

Dim linesPerpage As Single
linesPerpage = TickersContainerPicture.Height / mLineSpacing

If totalLines > linesPerpage Then mLinesToScroll = -Int(linesPerpage - totalLines)

Dim pagesToScroll As Single
pagesToScroll = mLinesToScroll / linesPerpage

If mLinesToScroll > 1 Then
    TickerScroll.SmallChange = (CLng(TickerScrollMax) + mLinesToScroll - 1) / mLinesToScroll
Else
    TickerScroll.SmallChange = TickerScrollMax
End If

If pagesToScroll > 1 Then
    TickerScroll.LargeChange = TickerScrollMax / pagesToScroll
Else
    TickerScroll.LargeChange = TickerScrollMax
End If
TickerScroll.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showActivityMonitor()
Const ProcName As String = "showActivityMonitor"
On Error GoTo Err

ShowHideMonitorButton.Caption = "Hide activity monitor"
mAdjustingSize = True
Me.Height = Me.Height + ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = True
mActivityMonitorVisible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startCollecting( _
                ByVal message As String)
Const ProcName As String = "startCollecting"
On Error GoTo Err

LogMessage message

mDataCollector.StartCollection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub stopCollecting( _
                ByVal message As String, _
                ByVal confirm As Boolean)
Const ProcName As String = "stopCollecting"
On Error GoTo Err

If confirm Then
    If MsgBox("Please confirm that you wish to stop data collection", _
                vbYesNo + vbDefaultButton2 + vbQuestion) <> vbYes Then Exit Sub
End If

LogMessage message

mDataCollector.StopCollection

Dim i As Long
For i = 0 To UBound(mTickers)
    If Not mTickers(i).Ticker Is Nothing Then
        Set mTickers(i).Ticker = Nothing
        switchDataLightOff i
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub switchDataLightOn( _
                ByVal index As Long)
Const ProcName As String = "switchDataLightOn"
On Error GoTo Err

If Not mActivityMonitorVisible Then Exit Sub

If Not mTickers(index).tli Is Nothing Then
    mTickers(index).tli.RemoveStateChangeListener Me
    mTickers(index).tli.Cancel
End If

Set mTickers(index).tli = mTimerList.Add(index, 200, ExpiryTimeUnitMilliseconds)
mTickers(index).tli.AddStateChangeListener Me

DataLightLabel(index).BackColor = vbGreen
DataLightLabel(index).Refresh
ConnectionStatusText.BackColor = vbGreen

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub switchDataLightOff( _
                ByVal index As Long)
Const ProcName As String = "switchDataLightOff"
On Error GoTo Err

If Not mTickers(index).tli Is Nothing Then
    mTickers(index).tli.RemoveStateChangeListener Me
    mTickers(index).tli.Cancel
    Set mTickers(index).tli = Nothing
End If

DataLightLabel(index).BackColor = vbButtonFace
DataLightLabel(index).Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

