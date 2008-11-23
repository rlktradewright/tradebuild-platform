VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#123.1#0"; "TradeBuildUI2-6.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form fChart3
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   1
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
         TabIndex        =   5
         Top             =   105
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   582
      End
      Begin MSComctlLib.Toolbar TimeframeToolbar 
         Height          =   330
         Left            =   10095
         TabIndex        =   3
         Top             =   105
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "timeframe"
               Object.ToolTipText     =   "Change the timeframe"
               Style           =   4
               Object.Width           =   2000
            EndProperty
         EndProperty
         Begin TradeBuildUI26.TimeframeSelector TimeframeSelector1 
            Height          =   330
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   582
         End
      End
      Begin MSComctlLib.Toolbar ChartToolsToolbar 
         Height          =   540
         Left            =   180
         TabIndex        =   2
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
   Begin TradeBuildUI26.TradeBuildChart TradeBuildChart1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   2760
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   7011
      ShowToobar      =   0   'False
      TwipsPerBar     =   100
   End
End
Attribute VB_Name = "fChart3"
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

Private Const ModuleName                        As String = "fChart3"

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

Private mBarTimePeriod As timePeriod

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
gSyncStudyPicker TradeBuildChart1.ChartManager, _
                "Study picker for " & mSymbol & _
                " (" & mBarTimePeriod.toString & ")"
End Sub

Private Sub Form_Load()
TradeBuildChart1.Top = ChartToolsToolbar.Height
TradeBuildChart1.Left = 0
TradeBuildChart1.Width = Me.ScaleWidth
TradeBuildChart1.Height = Me.ScaleHeight - CoolBar1.Height

TradeBuildChart1.updatePerTick = True

TimeframeSelector1.initialise
End Sub

Private Sub Form_Resize()
Resize
End Sub

Private Sub Form_Unload(cancel As Integer)
If mIsHistorical Then mTicker.stopTicker
Set mTicker = Nothing
gUnsyncStudyPicker
TradeBuildChart1.finish
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub ChartToolsToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case ChartToolsCommandStudies
    gShowStudyPicker TradeBuildChart1.ChartManager, _
                    mSymbol & _
                    " (" & mBarTimePeriod.toString & ")"
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

Private Sub TimeframeSelector1_Click()
TradeBuildChart1.ChangeTimeframe TimeframeSelector1.timeframeDesignator
End Sub

'================================================================================
' mController Event Handlers
'================================================================================

Private Sub mController_PointerModeChanged()
If mController.PointerMode = PointerModeSelection Then
    ChartToolsToolbar.buttons("selection").value = tbrPressed
Else
    ChartToolsToolbar.buttons("selection").value = tbrUnpressed
End If
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
If ev.state = TickerStates.TickerStateReady Then
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

Set mBarTimePeriod = chartspec.Timeframe
TimeframeSelector1.selectTimeframe mBarTimePeriod

TradeBuildChart1.showChart mTicker, chartspec

setCaption

ChartNavToolbar1.initialise TradeBuildChart1
Set mController = TradeBuildChart1.chartController
End Sub

Friend Sub showHistoricalChart( _
                ByVal pTicker As Ticker, _
                ByVal chartspec As ChartSpecifier, _
                ByVal fromtime As Date, _
                ByVal totime As Date)

mIsHistorical = True

Set mTicker = pTicker

Set mBarTimePeriod = chartspec.Timeframe
TimeframeSelector1.selectTimeframe mBarTimePeriod

TradeBuildChart1.showHistoricChart mTicker, chartspec, fromtime, totime

ChartNavToolbar1.initialise TradeBuildChart1
Set mController = TradeBuildChart1.chartController
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub createFibChartTool()
Dim ls As lineStyle
Dim tool As FibRetracementTool
Dim lineSpecs(4) As FibLineSpecifier

Set ls = TradeBuildChart1.chartController.DefaultLineStyle
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

Set tool = CreateFibRetracementTool(TradeBuildChart1.chartController, lineSpecs, LayerNumbers.LayerHighestUser)
Set mCurrentTool = tool
TradeBuildChart1.SetFocus
End Sub

Private Sub createLineChartTool()
Dim tool As LineTool
Dim ls As lineStyle

Set ls = TradeBuildChart1.chartController.DefaultLineStyle
ls.extended = True
ls.extendAfter = True
ls.includeInAutoscale = False

Set tool = CreateLineTool(TradeBuildChart1.chartController, ls, LayerBackground)
Set mCurrentTool = tool
TradeBuildChart1.SetFocus
End Sub

Private Sub Resize()
If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub
TradeBuildChart1.Width = Me.ScaleWidth
If Me.ScaleHeight >= CoolBar1.Height Then
    TradeBuildChart1.Height = Me.ScaleHeight - CoolBar1.Height
    TradeBuildChart1.Top = CoolBar1.Height
End If
End Sub

Private Sub setCaption()
Dim s As String

s = mSymbol & _
    " (" & mBarTimePeriod.toString & ")"
    
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

