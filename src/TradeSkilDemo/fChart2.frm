VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#114.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fChart2 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI26.TradeBuildChart TradeBuildChart1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   10610
      TwipsPerBar     =   100
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10530
      _ExtentX        =   18574
      _ExtentY        =   1005
      ButtonWidth     =   1111
      ButtonHeight    =   953
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Studies"
            Key             =   "studies"
            Object.ToolTipText     =   "Manage the studies displayed on the chart"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Lines"
            Key             =   "lines"
            Object.ToolTipText     =   "Draw lines"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Fib"
            Key             =   "fib"
            Object.ToolTipText     =   "Draw Fibonacci retracement lines"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "fChart2"
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

Private mBarTimePeriod As TimePeriod

Private mCurrentBid As String
Private mCurrentAsk As String
Private mCurrentTrade As String
Private mCurrentVolume As Long
Private mCurrentHigh As String
Private mCurrentLow As String
Private mPreviousClose As String

Private mIsHistorical As Boolean

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
TradeBuildChart1.Top = Toolbar1.Height
TradeBuildChart1.Left = 0
TradeBuildChart1.Width = Me.ScaleWidth
TradeBuildChart1.Height = Me.ScaleHeight - Toolbar1.Height

TradeBuildChart1.updatePerTick = True

End Sub

Private Sub Form_Resize()
If Me.ScaleWidth = 0 And _
    Me.ScaleHeight = 0 Then Exit Sub
TradeBuildChart1.Width = Me.ScaleWidth
If Me.ScaleHeight >= Toolbar1.Height Then
    TradeBuildChart1.Height = Me.ScaleHeight - Toolbar1.Height
End If
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "studies"
    gShowStudyPicker TradeBuildChart1.ChartManager, _
                    mSymbol & _
                    " (" & mBarTimePeriod.toString & ")"
Case "lines"
    createLineChartTool
Case "fib"
    createFibChartTool
End Select
End Sub

'================================================================================
' mTicker Interface Members
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
                ByVal initialNumberOfBars As Long, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal minimumTicksHeight As Long, _
                ByVal barTimePeriod As TimePeriod)

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

Set mBarTimePeriod = barTimePeriod

TradeBuildChart1.showChart mTicker, _
                        initialNumberOfBars, _
                        includeBarsOutsideSession, _
                        minimumTicksHeight, _
                        barTimePeriod

setCaption

End Sub

Friend Sub showHistoricalChart( _
                ByVal pTicker As Ticker, _
                ByVal initialNumberOfBars As Long, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal minimumTicksHeight As Long, _
                ByVal barTimePeriod As TimePeriod)

mIsHistorical = True

Set mTicker = pTicker

Set mBarTimePeriod = barTimePeriod

TradeBuildChart1.showHistoricChart mTicker, _
                        initialNumberOfBars, _
                        fromDate, _
                        toDate, _
                        includeBarsOutsideSession, _
                        minimumTicksHeight, _
                        barTimePeriod

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
ls.IncludeInAutoscale = False

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
ls.IncludeInAutoscale = False

Set tool = createLineTool(TradeBuildChart1.chartController, ls, LayerBackground)
Set mCurrentTool = tool
TradeBuildChart1.SetFocus
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
