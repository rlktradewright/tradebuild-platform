VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{41BEA792-C104-45F5-96C2-0BF81D749359}#1.0#0"; "TradeBuildUI.ocx"
Begin VB.Form fChart2 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI.TradeBuildChart TradeBuildChart1 
      Height          =   5175
      Left            =   360
      TabIndex        =   1
      Top             =   840
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   9128
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

Private WithEvents mTicker As TradeBuild.Ticker
Attribute mTicker.VB_VarHelpID = -1

Private mSymbol As String
Private mCurrentBid As String
Private mCurrentAsk As String
Private mCurrentTrade As String
Private mCurrentVolume As Long
Private mCurrentHigh As String
Private mCurrentLow As String
Private mPreviousClose As String

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Load()
TradeBuildChart1.Top = Toolbar1.Height
TradeBuildChart1.Left = 0
TradeBuildChart1.Width = Me.ScaleWidth
TradeBuildChart1.Height = Me.ScaleHeight - Toolbar1.Height

TradeBuildChart1.updatePerTick = True

End Sub

Private Sub Form_Resize()
TradeBuildChart1.Width = Me.ScaleWidth
TradeBuildChart1.Height = Me.ScaleHeight - Toolbar1.Height
End Sub

Private Sub Form_Unload(cancel As Integer)
Set mTicker = Nothing
TradeBuildChart1.finish
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "studies"
    TradeBuildChart1.showStudyPickerForm
End Select
End Sub

'================================================================================
' mTicker Interface Members
'================================================================================

Private Sub mTicker_ask(ev As TradeBuild.QuoteEvent)
mCurrentAsk = ev.priceString
setCaption
End Sub

Private Sub mTicker_bid(ev As TradeBuild.QuoteEvent)
mCurrentBid = ev.priceString
setCaption
End Sub

Private Sub mTicker_high(ev As TradeBuild.QuoteEvent)
mCurrentHigh = ev.priceString
setCaption
End Sub

Private Sub mTicker_Low(ev As TradeBuild.QuoteEvent)
mCurrentLow = ev.priceString
setCaption
End Sub

Private Sub mTicker_previousClose(ev As TradeBuild.QuoteEvent)
mPreviousClose = ev.priceString
setCaption
End Sub

Private Sub mTicker_trade(ev As TradeBuild.QuoteEvent)
mCurrentTrade = ev.priceString
setCaption
End Sub

Private Sub mTicker_volume(ev As TradeBuild.QuoteEvent)
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
                ByVal pTicker As TradeBuild.Ticker, _
                ByVal initialNumberOfBars As Long, _
                ByVal minimumTicksHeight As Long, _
                ByVal periodlength As Long, _
                ByVal periodUnits As TimePeriodUnits)


Set mTicker = pTicker
mSymbol = mTicker.Contract.specifier.localSymbol
mCurrentBid = mTicker.BidPriceString
mCurrentTrade = mTicker.TradePriceString
mCurrentAsk = mTicker.AskPriceString
mCurrentVolume = mTicker.Volume
mCurrentHigh = mTicker.highPriceString
mCurrentLow = mTicker.lowPriceString
mPreviousClose = mTicker.closePriceString

TradeBuildChart1.initialNumberOfBars = initialNumberOfBars
TradeBuildChart1.minimumTicksHeight = minimumTicksHeight
TradeBuildChart1.periodlength = periodlength
TradeBuildChart1.periodUnits = periodUnits

TradeBuildChart1.showChart mTicker

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub setCaption()
Me.caption = mSymbol & _
            "    B=" & mCurrentBid & _
            "  T=" & mCurrentTrade & _
            "  A=" & mCurrentAsk & _
            "  V=" & mCurrentVolume & _
            "  H=" & mCurrentHigh & _
            "  L=" & mCurrentLow & _
            "  C=" & mPreviousClose

End Sub
