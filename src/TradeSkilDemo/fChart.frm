VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#139.0#0"; "TradeBuildUI2-6.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form fChart 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI26.MultiChart MultiChart1 
      Height          =   5295
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   9340
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   953
      _CBWidth        =   12525
      _CBHeight       =   540
      _Version        =   "6.7.9782"
      Child1          =   "ChartToolsToolbar"
      MinWidth1       =   2880
      MinHeight1      =   540
      Width1          =   2880
      NewRow1         =   0   'False
      Child2          =   "ChartNavToolbar1"
      MinHeight2      =   330
      Width2          =   6705
      NewRow2         =   0   'False
      Child3          =   "TimeframeToolbar"
      MinHeight3      =   330
      Width3          =   2835
      NewRow3         =   0   'False
      Begin TradeBuildUI26.ChartNavToolbar ChartNavToolbar1 
         Height          =   330
         Left            =   3330
         TabIndex        =   4
         Top             =   105
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
      End
      Begin MSComctlLib.Toolbar TimeframeToolbar 
         Height          =   330
         Left            =   10095
         TabIndex        =   2
         Top             =   105
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "timeframe"
               Object.ToolTipText     =   "Change the timeframe for the current chart"
               Style           =   4
               Object.Width           =   1700
            EndProperty
         EndProperty
         Begin TradeBuildUI26.TimeframeSelector TimeframeSelector1 
            Height          =   330
            Left            =   0
            TabIndex        =   3
            ToolTipText     =   "Change the timeframe for the current chart"
            Top             =   0
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   582
         End
      End
      Begin MSComctlLib.Toolbar ChartToolsToolbar 
         Height          =   540
         Left            =   180
         TabIndex        =   1
         Top             =   0
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   953
         ButtonWidth     =   1111
         ButtonHeight    =   953
         Appearance      =   1
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Studies"
               Key             =   "studies"
               Object.ToolTipText     =   "Manage the studies displayed on the chart"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Select"
               Key             =   "selection"
               Description     =   "Select a chart object"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Lines"
               Key             =   "lines"
               Object.ToolTipText     =   "Draw lines"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Fib"
               Key             =   "fib"
               Object.ToolTipText     =   "Draw Fibonacci retracement lines"
            EndProperty
         EndProperty
      End
   End
End
Attribute VB_Name = "fChart"
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

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "fChart"

Private Const ChartToolsCommandStudies          As String = "studies"
Private Const ChartToolsCommandSelection        As String = "selection"
Private Const ChartToolsCommandLines            As String = "lines"
Private Const ChartToolsCommandFib              As String = "fib"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTicker As Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mSymbol As String

Private mCurrentBid As String
Private mCurrentAsk As String
Private mCurrentTrade As String
Private mCurrentVolume As Long
Private mCurrentHigh As String
Private mCurrentLow As String
Private mPreviousClose As String

Private mIsHistorical As Boolean

Private WithEvents mController As chartController
Attribute mController.VB_VarHelpID = -1

Private mCurrentTool As IChartTool

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Activate()
syncStudyPicker
End Sub

Private Sub Form_Load()
Resize
TimeframeSelector1.initialise
End Sub

Private Sub Form_Resize()
Resize
End Sub

Private Sub Form_Unload(cancel As Integer)
MultiChart1.Clear
If mIsHistorical Then mTicker.stopTicker
Set mTicker = Nothing
gUnsyncStudyPicker
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub ChartToolsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)

If MultiChart1.Count = 0 Then Exit Sub

Select Case Button.Key
Case ChartToolsCommandStudies
    gShowStudyPicker MultiChart1.ChartManager, _
                    mSymbol & _
                    " (" & MultiChart1.timePeriod.toString & ")"
Case ChartToolsCommandSelection
    setSelectionMode
Case ChartToolsCommandLines
    createLineChartTool
Case ChartToolsCommandFib
    createFibChartTool
End Select
End Sub

Private Sub CoolBar1_HeightChanged(ByVal NewHeight As Single)
Resize
End Sub

Private Sub MultiChart1_Change(ev As TWUtilities30.ChangeEvent)
Dim changeType As MultiChartChangeTypes
changeType = ev.changeType
Select Case changeType
Case MultiChartSelectionChanged
    If MultiChart1.Count > 0 Then
        ChartToolsToolbar.Enabled = True
        TimeframeSelector1.Enabled = True
        Set mController = MultiChart1.chartController
        setCaption
        setSelectionButton
        setTimeframeSelector
        syncStudyPicker
    Else
        setCaption
        ChartToolsToolbar.Enabled = False
        TimeframeSelector1.Enabled = False
        Set mController = Nothing
    End If
    Set mCurrentTool = Nothing
Case MultiChartAdd

Case MultiChartRemove
    gUnsyncStudyPicker
Case MultiChartTimeframeChanged
    If MultiChart1.Count > 0 Then Set mController = MultiChart1.chartController
    setCaption
    setSelectionButton
    setTimeframeSelector
    syncStudyPicker
End Select
End Sub

Private Sub TimeframeSelector1_Click()
If MultiChart1.Count > 0 Then
    MultiChart1.ChangeTimeframe TimeframeSelector1.timeframeDesignator
    setCaption
End If
End Sub

'================================================================================
' mController Event Handlers
'================================================================================

Private Sub mController_PointerModeChanged()
setSelectionButton
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_ask(ev As QuoteEvent)
mCurrentAsk = ev.priceString
setCaption
End Sub

Private Sub mTicker_bid(ev As QuoteEvent)
mCurrentBid = ev.priceString
setCaption
End Sub

Private Sub mTicker_high(ev As QuoteEvent)
mCurrentHigh = ev.priceString
setCaption
End Sub

Private Sub mTicker_Low(ev As QuoteEvent)
mCurrentLow = ev.priceString
setCaption
End Sub

Private Sub mTicker_previousClose(ev As QuoteEvent)
mPreviousClose = ev.priceString
setCaption
End Sub

Private Sub mTicker_stateChange( _
                ByRef ev As TWUtilities30.StateChangeEvent)
If ev.State = TickerStates.TickerStateReady Then
    mSymbol = mTicker.Contract.specifier.localSymbol
    setCaption
End If
End Sub

Private Sub mTicker_trade(ev As QuoteEvent)
mCurrentTrade = ev.priceString
setCaption
End Sub

Private Sub mTicker_volume(ev As QuoteEvent)
mCurrentVolume = ev.Size
setCaption
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub showChart( _
                ByVal pTicker As Ticker, _
                ByVal chartspec As ChartSpecifier)

mIsHistorical = False

Set mTicker = pTicker
mSymbol = mTicker.Contract.specifier.localSymbol
mCurrentBid = mTicker.BidPriceString
mCurrentTrade = mTicker.TradePriceString
mCurrentAsk = mTicker.AskPriceString
mCurrentVolume = mTicker.Volume
mCurrentHigh = mTicker.highPriceString
mCurrentLow = mTicker.lowPriceString
mPreviousClose = mTicker.closePriceString

TimeframeSelector1.selectTimeframe chartspec.Timeframe

MultiChart1.initialise mTicker, chartspec, , , New DunniganFactory
MultiChart1.Add chartspec.Timeframe

ChartNavToolbar1.initialise , MultiChart1

setCaption
End Sub

Friend Sub showHistoricalChart( _
                ByVal pTicker As Ticker, _
                ByVal chartspec As ChartSpecifier, _
                ByVal fromtime As Date, _
                ByVal totime As Date)

mIsHistorical = True

Set mTicker = pTicker

TimeframeSelector1.selectTimeframe chartspec.Timeframe

MultiChart1.initialise mTicker, chartspec, fromtime, totime, New DunniganFactory
MultiChart1.Add chartspec.Timeframe

ChartNavToolbar1.initialise , MultiChart1
setCaption
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub createFibChartTool()
Dim ls As lineStyle
Dim lineSpecs(4) As FibLineSpecifier

Set ls = mController.DefaultLineStyle
ls.extended = True
ls.includeInAutoscale = False

ls.Color = vbBlack
Set lineSpecs(0).Style = ls.Clone
lineSpecs(0).Percentage = 0

ls.Color = vbRed
Set lineSpecs(1).Style = ls.Clone
lineSpecs(1).Percentage = 100

ls.Color = &H8000&   ' dark green
Set lineSpecs(2).Style = ls.Clone
lineSpecs(2).Percentage = 50

ls.Color = vbBlue
Set lineSpecs(3).Style = ls.Clone
lineSpecs(3).Percentage = 38.2

ls.Color = vbMagenta
Set lineSpecs(4).Style = ls.Clone
lineSpecs(4).Percentage = 61.8

Set mCurrentTool = CreateFibRetracementTool(mController, lineSpecs, LayerNumbers.LayerHighestUser)
MultiChart1.SetFocus
End Sub

Private Sub createLineChartTool()
Dim ls As lineStyle

Set ls = mController.DefaultLineStyle
ls.extended = True
ls.extendAfter = True
ls.includeInAutoscale = False

Set mCurrentTool = CreateLineTool(mController, ls, LayerBackground)
MultiChart1.SetFocus
End Sub

Private Sub Resize()
If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub
MultiChart1.Width = Me.ScaleWidth
If Me.ScaleHeight >= CoolBar1.Height Then
    MultiChart1.Height = Me.ScaleHeight - CoolBar1.Height
    MultiChart1.Top = CoolBar1.Height
End If
End Sub

Private Sub setCaption()
Dim s As String

If MultiChart1.Count = 0 Then
    s = mSymbol
Else
    s = mSymbol & " (" & MultiChart1.timePeriod.toString & ")"
End If
    
If mIsHistorical Then
    s = s & _
        "    (historical)"
Else
    s = s & _
        "    B=" & mCurrentBid & _
        "  T=" & mCurrentTrade & _
        "  A=" & mCurrentAsk & _
        "  V=" & mCurrentVolume & _
        "  H=" & mCurrentHigh & _
        "  L=" & mCurrentLow & _
        "  C=" & mPreviousClose
End If
Me.caption = s
End Sub

Private Sub setSelectionMode()
If mController.PointerMode <> PointerModeSelection Then
    mController.SetPointerModeSelection
    ChartToolsToolbar.buttons("selection").value = tbrPressed
Else
    mController.SetPointerModeDefault
    ChartToolsToolbar.buttons("selection").value = tbrUnpressed
End If
End Sub

Private Sub setSelectionButton()
If mController.PointerMode = PointerModeSelection Then
    ChartToolsToolbar.buttons("selection").value = tbrPressed
Else
    ChartToolsToolbar.buttons("selection").value = tbrUnpressed
End If
End Sub

Private Sub setTimeframeSelector()
If MultiChart1.Count > 0 Then TimeframeSelector1.selectTimeframe MultiChart1.Object.Chart.timePeriod
End Sub

Private Sub syncStudyPicker()
If MultiChart1.Count = 0 Then Exit Sub
gSyncStudyPicker MultiChart1.ChartManager, _
                "Study picker for " & mSymbol & _
                " (" & MultiChart1.timePeriod.toString & ")"
End Sub

