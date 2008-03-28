VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#62.0#0"; "TradeBuildUI2-6.ocx"
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
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Studies"
            Key             =   "studies"
            Object.ToolTipText     =   "Manage the studies displayed on the chart"
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

Private mPeriodlength As Long
Private mPeriodUnits As TimePeriodUnits

Private mCurrentBid As String
Private mCurrentAsk As String
Private mCurrentTrade As String
Private mCurrentVolume As Long
Private mCurrentHigh As String
Private mCurrentLow As String
Private mPreviousClose As String

Private mIsHistorical As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Activate()
gSyncStudyPicker TradeBuildChart1.ChartManager, _
                "Study picker for " & mSymbol & _
                " (" & mPeriodlength & " " & TimeframeUtils.TimePeriodUnitsToString(mPeriodUnits) & ")"
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
If mIsHistorical Then mTicker.StopTicker
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
                    " (" & mPeriodlength & " " & TimeframeUtils.TimePeriodUnitsToString(mPeriodUnits) & ")"
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
                ByVal periodlength As Long, _
                ByVal periodUnits As TimePeriodUnits)

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

mPeriodlength = periodlength
mPeriodUnits = periodUnits

TradeBuildChart1.showChart mTicker, _
                        initialNumberOfBars, _
                        includeBarsOutsideSession, _
                        minimumTicksHeight, _
                        periodlength, _
                        periodUnits

setCaption

End Sub

Friend Sub showHistoricalChart( _
                ByVal pTicker As Ticker, _
                ByVal initialNumberOfBars As Long, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal includeBarsOutsideSession As Boolean, _
                ByVal minimumTicksHeight As Long, _
                ByVal periodlength As Long, _
                ByVal periodUnits As TimePeriodUnits)

mIsHistorical = True

Set mTicker = pTicker

mPeriodlength = periodlength
mPeriodUnits = periodUnits

TradeBuildChart1.showHistoricChart mTicker, _
                        initialNumberOfBars, _
                        fromDate, _
                        toDate, _
                        includeBarsOutsideSession, _
                        minimumTicksHeight, _
                        mPeriodlength, _
                        mPeriodUnits

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub setCaption()
Dim s As String

s = mSymbol & _
    " (" & mPeriodlength & " " & TimePeriodUnitsToString(mPeriodUnits) & ")"
    
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
