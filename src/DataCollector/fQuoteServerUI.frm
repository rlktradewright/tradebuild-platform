VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fDataCollectorUI 
   Caption         =   "TradeBuild Data Collector Version 2.6"
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
      Tabs            =   2
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
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
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

Implements BarWriterListener
Implements LogListener
Implements QuoteListener
Implements StateChangeListener
Implements TickfileWriterListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const TickerScrollMax As Integer = 32767
Private Const TickerScrollMin As Integer = 0

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type TickerTableEntry
    theTicker               As ticker
    tli                     As TimerListItem
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mDataCollector As dataCollector
Attribute mDataCollector.VB_VarHelpID = -1

Private mTickers() As TickerTableEntry

Private mTimerList As TimerList

Private mLastTickTime As Date

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mTickCount As Long

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

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
Set mTimerList = GetGlobalTimerList
ReDim mTickers(99) As TickerTableEntry
Set mFormatter = CreateBasicLogFormatter(TimestampTimeOnlyLocal)
gLogger.addLogListener Me
End Sub

Private Sub Form_Load()
mStartStopButtonInitialLeft = StartStopButton.Left
TickerScroll.Min = TickerScrollMin
TickerScroll.Max = TickerScrollMax
mLineSpacing = ShortNameText(0).Height - Screen.TwipsPerPixelY
mCurrentHeight = Me.Height
mCurrentWidth = Me.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
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
End Sub

Private Sub Form_Resize()

If Me.Height <> mCurrentHeight Then resizeHeight
If Me.Width <> mCurrentWidth Then resizeWidth

End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
gLogger.Log LogLevelNormal, "Data Collector program exiting"
gLogger.removeLogListener Me
End Sub

'================================================================================
' BarWriterListener Interface Members
'================================================================================

Private Sub BarWriterListener_notify(ev As TradeBuild26.WriterEvent)
Dim tk As ticker

Select Case ev.Action
Case WriterNotifications.WriterNotReady
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Bar writer not ready for " & _
                tk.Contract.specifier.localSymbol
Case WriterNotifications.WriterReady
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Bar writer ready for " & _
                tk.Contract.specifier.localSymbol
Case WriterNotifications.WriterFileCreated
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Writing bars for " & _
                tk.Contract.specifier.localSymbol & _
                " to " & ev.FileName
End Select
End Sub

'================================================================================
' LogListener Interface Members
'================================================================================

Private Sub LogListener_finish()
'nothing to do
End Sub

Private Sub LogListener_Notify(ByVal logrec As TWUtilities30.LogRecord)
LogText.SelStart = Len(LogText.Text)
LogText.SelLength = 0
If Len(LogText.Text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = mFormatter.formatRecord(logrec)
LogText.SelStart = InStrRev(LogText.Text, vbCrLf) + 2
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
processQuoteEvent ev
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
processQuoteEvent ev
End Sub

'================================================================================
' StateChangeListener Interface Members
'================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEvent)
Dim tli As TimerListItem

Set tli = ev.Source
If Not tli Is Nothing Then
    If ev.State = TimerListItemStates.TimerListItemStateExpired Then
        switchDataLightOff tli.Data
        Set mTickers(tli.Data).tli = Nothing
    End If
End If
End Sub

'================================================================================
' TickfileWriterListener Interface Members
'================================================================================

Private Sub TickfileWriterListener_notify(ev As TradeBuild26.WriterEvent)
Dim tk As ticker

Select Case ev.Action
Case WriterNotifications.WriterNotReady
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Tickfile writer not ready for " & _
                    tk.Contract.specifier.localSymbol
Case WriterNotifications.WriterReady
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Tickfile writer ready for " & _
                    tk.Contract.specifier.localSymbol
Case WriterNotifications.WriterFileCreated
    Set tk = ev.Source
    gLogger.Log LogLevelNormal, "Writing tickdata for " & _
                tk.Contract.specifier.localSymbol & _
                " to " & ev.FileName
End Select
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub ShowHideMonitorButton_Click()
If mActivityMonitorVisible Then
    hideActivityMonitor
Else
    showActivityMonitor
End If
End Sub

Private Sub StartStopButton_Click()
If mCollectingData Then
    stopCollecting "Data collection stopped by user"
Else
    startCollecting "Data collection started by user"
End If
End Sub

Private Sub TickerScroll_Change()
scrollTickers
End Sub

'================================================================================
' mDataCollector Event Handlers
'================================================================================

Private Sub mDataCollector_CollectionStarted()
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
End Sub

Private Sub mDataCollector_CollectionStopped()
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
End Sub

Private Sub mDataCollector_connected()
ConnectionStatusText.BackColor = vbGreen
StartStopButton.Enabled = True
End Sub

Private Sub mDataCollector_connectFailed(ByVal description As String)
ConnectionStatusText.BackColor = vbRed
End Sub

Private Sub mDataCollector_ConnectionClosed()
ConnectionStatusText.BackColor = vbRed
StartStopButton.Enabled = False
End Sub

Private Sub mDataCollector_ExitProgram()
Unload Me
End Sub

Private Sub mDataCollector_Reconnecting()
StartStopButton.Enabled = True
End Sub

Private Sub mDataCollector_TickerAdded(ByVal ticker As ticker)
Dim i As Long
Dim index As Long

index = ticker.Handle
If index > UBound(mTickers) Then
    ReDim Preserve mTickers(index / 100 * 100 + 99) As TickerTableEntry
End If
Set mTickers(index).theTicker = ticker
Set mTickers(index).tli = Nothing

If index > ShortNameText.UBound Then
    For i = ShortNameText.UBound + 1 To index
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

ShortNameText(index) = ticker.Contract.specifier.localSymbol
ShortNameText(index).ToolTipText = ticker.Contract.specifier.localSymbol
ShortNameText(index).Visible = True
DataLightLabel(index).Visible = True

ticker.addQuoteListener Me
ticker.addTickfileWriterListener Me
ticker.addBarWriterListener Me

Me.Refresh

End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Static timerCount As Long
timerCount = timerCount + 1
If timerCount Mod 4 = 0 Then
    TicksPerSecText = mTickCount
    mTickCount = 0
End If
SecsSinceLastTickText = Format(86400 * (GetTimestamp - mLastTickTime), "0")
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let dataCollector(ByVal value As TBDataCollector)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal pDataCollector As dataCollector, _
                ByVal noAutoStart As Boolean, _
                ByVal showMonitor As Boolean)

Set mStartStopTimePanel = StatusBar1.Panels.Item(1)

Set mDataCollector = pDataCollector

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

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearTickers()
Dim i As Long

For i = 0 To ShortNameText.UBound
    ShortNameText(i).Text = ""
    DataLightLabel(i).BackColor = vbButtonFace
Next
End Sub

Private Sub hideActivityMonitor()
ShowHideMonitorButton.Caption = "Show activity monitor"
mAdjustingSize = True
Me.Height = Me.Height - ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = False
mActivityMonitorVisible = False
End Sub
    
Private Sub processQuoteEvent( _
                ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
switchDataLightOn lTicker.Handle
mLastTickTime = GetTimestamp
mTickCount = mTickCount + 1
End Sub

Private Sub resizeHeight()
Dim heightIncrement As Long

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Height = mCurrentHeight
    Exit Sub
End If

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
End Sub

Private Sub resizeWidth()
Dim widthIncrement As Long

If Not mActivityMonitorVisible And Not mAdjustingSize Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

If Me.Width <= mStartStopButtonInitialLeft + StartStopButton.Width + 120 Then
    Me.Width = mCurrentWidth
    Exit Sub
End If

widthIncrement = Me.Width - mCurrentWidth

If Not mAdjustingSize Then
    
    ActivityMonitor.Width = ActivityMonitor.Width + widthIncrement

    TickersPicture.Width = TickersPicture.Width + widthIncrement

    TickerScroll.Left = TickerScroll.Left + widthIncrement

    LogText.Width = LogText.Width + widthIncrement

    StartStopButton.Left = StartStopButton.Left + widthIncrement
End If

mCurrentWidth = Me.Width

End Sub

Private Sub scrollTickers()
If TickersPicture.Height <= TickersContainerPicture.Height Then
    TickersPicture.Top = 0
ElseIf TickerScroll.value = TickerScrollMax Then
    TickersPicture.Top = -mLinesToScroll * mLineSpacing
Else
    TickersPicture.Top = -Round((mLinesToScroll / TickerScrollMax) * TickerScroll.value, 0) * mLineSpacing
End If
End Sub

Private Sub setStarted()
mCollectingData = True
StartStopButton.Caption = "Stop"
StartStopButton.Enabled = True

mLastTickTime = GetTimestamp

Set mTimer = CreateIntervalTimer(0, , 250)
mTimer.StartTimer
End Sub

Private Sub setStopped()
mCollectingData = False
StartStopButton.Caption = "Start"

ConnectionStatusText.BackColor = vbButtonFace

clearTickers

SecsSinceLastTickText = ""
TicksPerSecText = ""

mTimer.StopTimer
End Sub

Private Sub setupTickerScroll()
Dim totalLines As Long
Dim linesPerpage As Single
Dim pagesToScroll As Single

totalLines = (ShortNameText.UBound + 5) / 5
linesPerpage = TickersContainerPicture.Height / mLineSpacing
If totalLines > linesPerpage Then mLinesToScroll = -Int(linesPerpage - totalLines)
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
End Sub

Private Sub showActivityMonitor()
ShowHideMonitorButton.Caption = "Hide activity monitor"
mAdjustingSize = True
Me.Height = Me.Height + ActivityMonitor.Height
mAdjustingSize = False
ActivityMonitor.Visible = True
mActivityMonitorVisible = True
End Sub

Private Sub startCollecting( _
                ByVal message As String)
                
gLogger.Log LogLevelNormal, message

mDataCollector.startCollection

End Sub

Private Sub stopCollecting( _
                ByVal message As String)
If MsgBox("Please confirm that you wish to stop data collection", _
            vbYesNo + vbDefaultButton2 + vbQuestion) <> vbYes Then Exit Sub

gLogger.Log LogLevelNormal, message

mDataCollector.stopCollection

End Sub

Private Sub switchDataLightOn( _
                ByVal index As Long)
If Not mActivityMonitorVisible Then Exit Sub

If Not mTickers(index).tli Is Nothing Then
    mTimerList.Remove mTickers(index).tli
    mTickers(index).tli.removeStateChangeListener Me
End If

Set mTickers(index).tli = mTimerList.Add(index, 200, ExpiryTimeUnitMilliseconds)
mTickers(index).tli.addStateChangeListener Me

DataLightLabel(index).BackColor = vbGreen
ConnectionStatusText.BackColor = vbGreen
End Sub

Private Sub switchDataLightOff( _
                ByVal index As Long)
DataLightLabel(index).BackColor = vbButtonFace
End Sub

