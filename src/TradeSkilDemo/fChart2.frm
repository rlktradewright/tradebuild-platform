VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9D2C4B5E-2539-4900-8B70-B9B41CFF1CA8}#18.0#0"; "TradeBuildUI2-5.ocx"
Begin VB.Form fChart2 
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6780
   ScaleWidth      =   10530
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI25.TradeBuildChart TradeBuildChart1 
      Height          =   6015
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10610
      PointerDiscColor=   13893631
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

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Activate()
TradeBuildChart1.syncStudyPickerForm
End Sub

Private Sub Form_Deactivate()
TradeBuildChart1.unsyncStudyPickerForm
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
Set mTicker = Nothing
TradeBuildChart1.unsyncStudyPickerForm
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

'================================================================================
' Helper Functions
'================================================================================

Private Sub setCaption()
Me.caption = mSymbol & _
            " (" & mPeriodlength & " " & TimePeriodUnitsToString(mPeriodUnits) & ")" & _
            "    B=" & mCurrentBid & _
            "  T=" & mCurrentTrade & _
            "  A=" & mCurrentAsk & _
            "  V=" & mCurrentVolume & _
            "  H=" & mCurrentHigh & _
            "  L=" & mCurrentLow & _
            "  C=" & mPreviousClose

End Sub
