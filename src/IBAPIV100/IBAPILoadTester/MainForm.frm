VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "API  Load Tester V100"
   ClientHeight    =   8055
   ClientLeft      =   360
   ClientTop       =   435
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   9240
   Begin VB.TextBox ClientIdText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "645326819"
      Top             =   840
      Width           =   1095
   End
   Begin VB.TextBox PortText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1800
      TabIndex        =   1
      Text            =   "7497"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox ServerText 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton ConnectButton 
      Caption         =   "Connect"
      Default         =   -1  'True
      Height          =   495
      Left            =   7200
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.TextBox MaxCpuPercentPerSecText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox MaxEventsPerSecText 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   2880
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Last second"
      Height          =   1095
      Left            =   3120
      TabIndex        =   18
      Top             =   1560
      Width           =   2535
      Begin VB.TextBox CpuUtilisationThisSecondText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox ProcessTimeLastSecondText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox EventsLastSecondText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "CPU %"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "Process time"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Events"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Overall"
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   2895
      Begin VB.TextBox MicrosecsPerTickText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox AvgCpuUtilisationText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TotalProcessTimeText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox SecondsElapsedText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox AvgEventsPerSecondText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TotalEventsText 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Microsecs per tick"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Avg CPU %"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label8 
         Caption         =   "Total process time"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Seconds elapsed"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Avg events per sec"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Total events"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton StopTickCountingButton 
      Caption         =   "Stop counting ticks"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton StartTickCountingButton 
      Caption         =   "Start counting ticks"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton StopTickersButton 
      Caption         =   "Stop tickers"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7200
      TabIndex        =   6
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox LogText 
      Height          =   4335
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3600
      Width           =   9015
   End
   Begin VB.CommandButton StartTickersButton 
      Caption         =   "Start tickers"
      Height          =   495
      Left            =   7200
      TabIndex        =   4
      Top             =   720
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "To specify the tickers, edit the Symbols.txt file in the Data sub-folder."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Index           =   1
      Left            =   7200
      TabIndex        =   36
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label14 
      Caption         =   "Client Id"
      Height          =   375
      Left            =   360
      TabIndex        =   35
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label13 
      Caption         =   "Port"
      Height          =   375
      Left            =   360
      TabIndex        =   34
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label12 
      Caption         =   "TWS Server"
      Height          =   375
      Left            =   360
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label10 
      Caption         =   "Max CPU % per sec"
      Height          =   255
      Left            =   3120
      TabIndex        =   30
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Max events per sec"
      Height          =   255
      Left            =   3120
      TabIndex        =   28
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================

' A simple program that uses the TradeWright TWS API to control a set of tickers,
' and keep track of the number of ticks received via the API, and the
' CPU utilisation.
'
' The symbols to use are specified in the symbols.txt file which must be in the
' DATA subfolder below the folder that the program is running in. The symbols.txt
' file supplied in the download contains a number of busy tickers, but you can
' easily edit it to include whatever you like.

'================================================================================
' Interfaces
'================================================================================

Implements IConnectionStatusConsumer
Implements IErrorAndNotificationConsumer
Implements IMarketDataConsumer
Implements IMarketDepthConsumer

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External function declarations
'================================================================================

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function GetProcessTimes Lib "kernel32" ( _
                ByVal hProcess As Long, _
                lpCreationTime As Currency, _
                lpExitTime As Currency, _
                lpKernelTime As Currency, _
                lpUserTime As Currency) As Long
                
'================================================================================
' Member variables
'================================================================================

Private mCounting               As Boolean  ' Set to True while counting ticks
                                            ' received from the API

Private mTotalTicks             As Long     ' total number of tick events via API

Private mTicksThisSecond        As Long     ' number of tick events in the current second

Private mMaxEventsPerSecond     As Long     ' maximum number of tick events received in
                                            ' a second

Private mMaxCpuPerSecond        As Currency ' maximum CPU time in a second

Private mSecondsSinceStart      As Long     ' number of seconds since the start
                                            ' button was pressed

Private mNextTickerId           As Long     ' The id to use for the next market data ticker
Private mNextDepthTickerId      As Long     ' The id to use for the next market depth ticker

Private mProcessHandle          As Long     ' handle to this process allowing us to
                                            ' get info about the process from
                                            ' Windows
                                            
' NB: the following variables are of type Currency, because this type is actually
' held internally as a 64-bit integer, which is what is expected by the
' GetProcessTimes API function. Note that Currency variables are scaled down by 10,000
' when used: eg the value 1 held in the internal 64-bit number is interpreted as
' 0.0001. Since the values returned by GetProcessTimes are in 100-nanosecond units,
' this means that we can use the value in the Currency variable directly as a
' number of milliseconds.
Private mInitialCPUTime         As Currency ' the amount of CPU time the process had
                                            ' spent at start of tick counting
Private mPrevCPUTime            As Currency ' the amount of CPU time the process had
                                            ' spent at the end of the previous second
                                            
Private WithEvents mPerformanceTimer As IntervalTimer
Attribute mPerformanceTimer.VB_VarHelpID = -1

Private mIBAPI                  As TwsAPI

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Load()
InitialiseCommonControls
ApplicationGroupName = "TradeWright"
ApplicationName = "ApiLoadTesterV100"

SetupDefaultLogging Command

mProcessHandle = GetCurrentProcess

End Sub

Private Sub Form_Unload(Cancel As Integer)
If mIBAPI Is Nothing Then Exit Sub
If mIBAPI.ConnectionState = TwsConnConnected Then
    logMessage "Disconnecting from TWS"
    mIBAPI.Disconnect "Form closed"
End If
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub ClientIdText_GotFocus()
ClientIdText.SelStart = 0
ClientIdText.SelLength = Len(ClientIdText.Text)
End Sub

Private Sub ConnectButton_Click()
If ConnectButton.Caption = "Connect" Then
    Set mIBAPI = GetAPI(ServerText, CLng(PortText), CLng(ClientIdText), pLogApiMessageStats:=True)
    mIBAPI.ConnectionStatusConsumer = Me
    mIBAPI.ErrorAndNotificationConsumer = Me
    mIBAPI.MarketDataConsumer = Me
    ConnectButton.Enabled = False
    ConnectButton.MousePointer = vbHourglass
    mIBAPI.Connect
Else
    logMessage "Disconnecting from TWS"
    mIBAPI.Disconnect "User disconnected"
    ConnectButton.Caption = "Connect"
End If
End Sub

Private Sub PortText_GotFocus()
PortText.SelStart = 0
PortText.SelLength = Len(PortText.Text)
End Sub

Private Sub ServerText_GotFocus()
ServerText.SelStart = 0
ServerText.SelLength = Len(ServerText.Text)
End Sub

Private Sub StartTickersButton_Click()
mNextTickerId = 0

clearPerformanceFields

startTickers

StartTickersButton.Enabled = False
StopTickersButton.Enabled = True

StartTickCountingButton.Enabled = True
StartTickCountingButton.SetFocus
End Sub

Private Sub StartTickCountingButton_Click()
clearPerformanceFields

StartTickCountingButton.Enabled = False
StopTickCountingButton.Enabled = True
StopTickCountingButton.SetFocus

mCounting = True
End Sub

Private Sub StopTickersButton_Click()
Dim i As Long

If Not mPerformanceTimer Is Nothing Then mPerformanceTimer.StopTimer

For i = 0 To mNextTickerId - 1
    mIBAPI.CancelMarketData i
    logMessage "Stopping market data ticker " & i
Next
mNextTickerId = 0

For i = 0 To mNextDepthTickerId - 1
    mIBAPI.CancelMarketDepth i
    logMessage "Stopping market depth ticker " & i
Next
mNextDepthTickerId = 0

StartTickersButton.Enabled = True
StartTickersButton.SetFocus
StopTickersButton.Enabled = False

StartTickCountingButton.Enabled = False
StopTickCountingButton.Enabled = False

mCounting = False
StopTickCountingButton.Enabled = False
End Sub

Private Sub StopTickCountingButton_Click()
StartTickCountingButton.Enabled = True
StartTickCountingButton.SetFocus
StopTickCountingButton.Enabled = False
mCounting = False
If Not mPerformanceTimer Is Nothing Then mPerformanceTimer.StopTimer
logMessage "Tick counting stopped"

End Sub

'================================================================================
' IConnectionStatusConsumer Members
'================================================================================

Private Sub IConnectionStatusConsumer_NotifyAPIConnectionStateChange( _
                ByVal pState As IBAPIV100.TwsConnectionStates, _
                ByVal pMessage As String)
Select Case pState
Case TwsConnNotConnected
    logMessage "Disconnected from TWS"
    
    ConnectButton.Enabled = True
    ConnectButton.MousePointer = vbDefault
    ConnectButton.Caption = "Connect"
    ConnectButton.SetFocus
Case TwsConnConnecting
    logMessage "Connecting to TWS"
Case TwsConnConnected
    logMessage "Connected to TWS"
    
    ConnectButton.Enabled = True
    ConnectButton.MousePointer = vbDefault
    ConnectButton.Caption = "Disconnect"
    
    StartTickersButton.Enabled = True
    StartTickersButton.SetFocus
Case TwsConnFailed
    logMessage "Can't connect to TWS"
    
    ConnectButton.Enabled = True
    ConnectButton.MousePointer = vbDefault
    ConnectButton.Caption = "Connect"
    ConnectButton.SetFocus
End Select
End Sub

Private Sub IConnectionStatusConsumer_NotifyIBServerConnectionClosed()

End Sub

Private Sub IConnectionStatusConsumer_NotifyIBServerConnectionRecovered(ByVal pDataLost As Boolean)

End Sub

'================================================================================
' IErrorAndNotificationConsumer Members
'================================================================================

Private Sub IErrorAndNotificationConsumer_NotifyApiError(ByVal pErrorCode As Long, ByVal pErrorMsg As String)
logMessage "Error " & ": " & pErrorMsg
End Sub

Private Sub IErrorAndNotificationConsumer_NotifyApiEvent(ByVal pEventCode As Long, ByVal pEventMsg As String)
logMessage "Event " & ": " & pEventMsg
End Sub

'================================================================================
' IMarketDataConsumer Members
'================================================================================

Private Sub IMarketDataConsumer_EndTickSnapshot(ByVal pReqId As Long)

End Sub

Private Sub IMarketDataConsumer_NotifyError(ByVal pTickerId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
logMessage "Market data error " & pErrorCode & "(" & pTickerId & "): " & pErrorMsg
End Sub

Private Sub IMarketDataConsumer_NotifyTickEFP(ByVal pTickerId As Long, ByVal pTickType As IBAPIV100.TwsTickTypes, ByVal pBasisPoints As Double, ByVal pFormattedBasisPoints As String, ByVal pTotalDividends As Double, ByVal pHoldDays As Long, ByVal pFutureExpiry As String, ByVal pDividendImpact As Double, ByVal pDividendsToExpiry As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickGeneric(ByVal pTickerId As Long, ByVal pTickType As IBAPIV100.TwsTickTypes, ByVal pValue As Double)
incrementTotalTicks
End Sub

Private Sub IMarketDataConsumer_NotifyTickOptionComputation(ByVal pTickerId As Long, ByVal pTickType As IBAPIV100.TwsTickTypes, ByVal pImpliedVol As Double, ByVal pDelta As Double, ByVal pOptPrice As Double, ByVal pPvDividend As Double, ByVal pGamma As Double, ByVal pVega As Double, ByVal pTheta As Double, ByVal pUndPrice As Double)

End Sub

Private Sub IMarketDataConsumer_NotifyTickPrice(ByVal pTickerId As Long, ByVal pTickType As IBAPIV100.TwsTickTypes, ByVal pPrice As Double, ByVal pSize As Long, ByRef pAttributes As TwsTickAttributes)
incrementTotalTicks
End Sub

Private Sub IMarketDataConsumer_NotifyTickRequestParams(ByVal pTickerId As Long, ByVal pMinTick As Double, ByVal pBboExchange As String, ByVal pSnapshotPermissions As Long)
incrementTotalTicks
End Sub

Private Sub IMarketDataConsumer_NotifyTickSize(ByVal pTickerId As Long, ByVal pTickType As Long, ByVal pSize As Long)
incrementTotalTicks
End Sub

Private Sub IMarketDataConsumer_NotifyTickString(ByVal pTickerId As Long, ByVal pTickType As IBAPIV100.TwsTickTypes, ByVal pValue As String)
incrementTotalTicks
End Sub

'================================================================================
' IMarketDepthConsumer Members
'================================================================================

Private Sub IMarketDepthConsumer_NotifyError(ByVal pTickerId As Long, ByVal pErrorCode As Long, ByVal pErrorMsg As String)
logMessage "Market depth error " & pErrorCode & "(" & pTickerId & "): " & pErrorMsg
End Sub

Private Sub IMarketDepthConsumer_NotifyMarketDepth(ByVal pTickerId As Long, ByVal pPosition As Long, ByVal pMarketMaker As String, ByVal pOperation As IBAPIV100.TwsDOMOperations, ByVal pSide As IBAPIV100.TwsDOMSides, ByVal pPrice As Double, ByVal pSize As Long)
incrementTotalTicks
End Sub

Private Sub IMarketDepthConsumer_ResetMarketDepth(ByVal pReEstablish As Boolean)

End Sub

'================================================================================
' mPerformanceTimer Event Handlers
'================================================================================

Private Sub mPerformanceTimer_TimerExpired(ev As TimerExpiredEventData)

mSecondsSinceStart = mSecondsSinceStart + 1
SecondsElapsedText.Text = mSecondsSinceStart
TotalEventsText.Text = mTotalTicks
AvgEventsPerSecondText.Text = Format(mTotalTicks / mSecondsSinceStart, "0")
EventsLastSecondText.Text = mTicksThisSecond

If mTicksThisSecond > mMaxEventsPerSecond Then
    mMaxEventsPerSecond = mTicksThisSecond
    MaxEventsPerSecText.Text = mMaxEventsPerSecond
End If

Dim cpuTime As Currency: cpuTime = getCpuTime
If cpuTime <> mPrevCPUTime Then
    TotalProcessTimeText.Text = Format(cpuTime - mInitialCPUTime, "0.000")
    AvgCpuUtilisationText.Text = Format((cpuTime - mInitialCPUTime) / 1000 / mSecondsSinceStart, "0.00%")
    MicrosecsPerTickText.Text = Format(1000 * (cpuTime - mInitialCPUTime) / mTotalTicks, "0.00")
End If

Dim cpuThisSecond As Currency: cpuThisSecond = cpuTime - mPrevCPUTime

ProcessTimeLastSecondText.Text = Format(cpuThisSecond, "0.000")
CpuUtilisationThisSecondText.Text = Format(cpuThisSecond / 1000, "0.00%")

If cpuThisSecond > mMaxCpuPerSecond Then
    mMaxCpuPerSecond = cpuThisSecond
    MaxCpuPercentPerSecText.Text = Format(mMaxCpuPerSecond / 1000, "0.00%")
End If

mTicksThisSecond = 0
mPrevCPUTime = getCpuTime
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub clearPerformanceFields()
mTotalTicks = 0
mTicksThisSecond = 0
mSecondsSinceStart = 0
mMaxEventsPerSecond = 0
mMaxCpuPerSecond = 0

TotalEventsText.Text = ""
EventsLastSecondText.Text = ""
AvgEventsPerSecondText.Text = ""
SecondsElapsedText.Text = ""
TotalProcessTimeText.Text = ""
AvgCpuUtilisationText.Text = ""
ProcessTimeLastSecondText.Text = ""
CpuUtilisationThisSecondText.Text = ""
MaxEventsPerSecText.Text = ""
MaxCpuPercentPerSecText.Text = ""
End Sub

' returns the total number of seconds of CPU time used by the process
Private Function getCpuTime() As Currency
Dim creationTime As Currency
Dim exitTime As Currency
Dim kernelTime As Currency
Dim userTime As Currency

GetProcessTimes mProcessHandle, creationTime, exitTime, kernelTime, userTime
getCpuTime = kernelTime + userTime
End Function

Private Function getToken(ByRef tokens() As String, ByVal index As Long) As String
If UBound(tokens) >= index Then
    getToken = tokens(index)
Else
    getToken = ""
End If
End Function

Private Sub incrementTotalTicks()
If mTotalTicks = 0 And mCounting Then
    Set mPerformanceTimer = startPerformanceTimer()
    mInitialCPUTime = getCpuTime()
    mPrevCPUTime = mInitialCPUTime

    logMessage ("Tick counting started")
End If

If mCounting Then
    mTotalTicks = mTotalTicks + 1
    mTicksThisSecond = mTicksThisSecond + 1
End If
End Sub

Private Sub logMessage( _
                ByVal message As String)
Dim l As Long

l = Len(LogText.Text)

LogText.SelStart = l
LogText.SelLength = 0

If l <> 0 Then LogText.SelText = vbCrLf
LogText.SelText = FormatDateTime(Now, vbLongTime)
LogText.SelText = "  "
LogText.SelText = message
End Sub

Private Function startPerformanceTimer() As IntervalTimer
Dim performanceTimer As IntervalTimer
Set performanceTimer = CreateIntervalTimer(1000, , 1000)
performanceTimer.StartTimer
Set startPerformanceTimer = performanceTimer
End Function

Private Sub startTicker(ByVal symbol As String, _
                ByVal sectype As String, _
                ByVal expiry As String, _
                ByVal exchange As String, _
                ByVal currencyCode As String, _
                Optional ByVal primaryExchange As String, _
                Optional ByVal multiplier As String, _
                Optional marketDepth As String)
Dim lContract As New TwsContractSpecifier
With lContract
    .symbol = symbol
    .sectype = TwsSecTypeFromString(sectype)
    .expiry = expiry
    .exchange = exchange
    .currencyCode = currencyCode
    If multiplier = "" Then
        .multiplier = 1
    ElseIf CInt(multiplier) = 0 Then
        .multiplier = 1
    Else
        .multiplier = CInt(multiplier)
    End If
    .PrimaryExch = primaryExchange
End With

logMessage "Starting market data ticker for " & symbol & ": id=" & mNextTickerId
mIBAPI.RequestMarketData mNextTickerId, lContract, "", False, False
mNextTickerId = mNextTickerId + 1

If marketDepth = "" Then
ElseIf CBool(marketDepth) Then
    logMessage "Starting market depth ticker for " & symbol & ": id=" & mNextDepthTickerId
    mIBAPI.RequestMarketDepth mNextDepthTickerId, lContract
    mNextDepthTickerId = mNextDepthTickerId + 1
End If

End Sub

Private Sub startTickers()
Dim fs As FileSystemObject
Dim inFile As TextStream
Dim inBuff As String
Dim tokens() As String
Dim lineNum As Long

On Error GoTo Err

Set fs = New FileSystemObject
Set inFile = fs.OpenTextFile(App.Path & "\Data\symbols.txt", ForReading)

Do While Not inFile.AtEndOfStream
    inBuff = UCase$(inFile.ReadLine)
    lineNum = lineNum + 1
    
    'ignore comment lines
    Do While inBuff = "" Or Left$(inBuff, 2) = "//"
        If inFile.AtEndOfStream Then Exit Sub
        inBuff = inFile.ReadLine
    Loop
    
    If inBuff = "EXIT" Then Exit Do
    
    inBuff = Replace(inBuff, " ", "")
    inBuff = Replace(inBuff, vbTab, "")
    
    tokens = Split(inBuff, ",")
    If UBound(tokens) < 2 Then
        logMessage "Missing column(s) in line " & lineNum & " of symbols.txt"
    Else

        startTicker getToken(tokens, 0), _
                    getToken(tokens, 1), _
                    getToken(tokens, 2), _
                    getToken(tokens, 3), _
                    getToken(tokens, 4), _
                    getToken(tokens, 5), _
                    getToken(tokens, 6), _
                    getToken(tokens, 7)
    End If
Loop

Exit Sub

Err:
logMessage "Unspecified problem with symbols.txt: " & Err.Description
End Sub


