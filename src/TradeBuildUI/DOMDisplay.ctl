VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#30.0#0"; "TWControls10.ocx"
Begin VB.UserControl DOMDisplay 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   1725
   ScaleWidth      =   3795
   Begin TWControls10.TWGrid DOMGrid 
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   2566
   End
End
Attribute VB_Name = "DOMDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements MarketDepthListener

'@================================================================================
' Events
'@================================================================================

Event Halted()
Event Resumed()

'@================================================================================
' Constants
'@================================================================================

Private Const ScrollbarWidth As Long = 370  ' value discovered by trial and error!

'@================================================================================
' Enums
'@================================================================================

Private Enum DOMColumns
    PriceLeft
    bidSize
    LastSize
    AskSize
    PriceRight
End Enum

Private Enum GridColours
    BGDefault = &HE1F4FD
    BGBid = &HD6C6F2
    BGAsk = &HFAE968
    BGLast = &HC1F7CA
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTicker As ticker
Attribute mTicker.VB_VarHelpID = -1

Private mContract As Contract
Private mInitialPrice As Double
Private mPriceIncrement As Double
Private mNumberOfVisibleRows As Long
Private mBasePrice As Double
Private mCeilingPrice As Double

Private mAskPrices() As Double
Private mMaxAskPricesIndex As Long
Private mBidPrices() As Double
Private mMaxBidPricesIndex As Long

Private mCurrentLast As Double

Private mHalted As Boolean

Private WithEvents mCentreTimer As IntervalTimer
Attribute mCentreTimer.VB_VarHelpID = -1

Private mIsVisible As Boolean

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Hide()
mIsVisible = False
End Sub

Private Sub UserControl_Initialize()

ReDim mAskPrices(20) As Double
ReDim mBidPrices(20) As Double

DOMGrid.AllowUserResizing = TWControls10.AllowUserResizeSettings.TwGridResizeColumns
DOMGrid.BackColorFixed = vbButtonFace
DOMGrid.BackColorSel = GridColours.BGDefault
DOMGrid.backColor = GridColours.BGDefault
DOMGrid.ForeColorSel = DOMGrid.foreColor
DOMGrid.Rows = 50
DOMGrid.Cols = 5
DOMGrid.FixedCols = 0
DOMGrid.FixedRows = 1
DOMGrid.FocusRect = TWControls10.FocusRectSettings.TwGridFocusNone
DOMGrid.ScrollBars = TWControls10.ScrollBarsSettings.TwGridScrollBarVertical



setCellContents 0, DOMColumns.PriceLeft, "Price"
DOMGrid.colWidth(DOMColumns.PriceLeft) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceLeft) = TWControls10.AlignmentSettings.TwGridAlignRightCenter

setCellContents 0, DOMColumns.bidSize, "Bids"
DOMGrid.colWidth(DOMColumns.bidSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.bidSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.LastSize, "Last"
DOMGrid.colWidth(DOMColumns.LastSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.LastSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.AskSize, "Asks"
DOMGrid.colWidth(DOMColumns.AskSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.AskSize) = TWControls10.AlignmentSettings.TwGridAlignCenterCenter

setCellContents 0, DOMColumns.PriceRight, "Price"
DOMGrid.colWidth(DOMColumns.PriceRight) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceRight) = TWControls10.AlignmentSettings.TwGridAlignLeftCenter

End Sub

Private Sub UserControl_Resize()
resize
End Sub

Private Sub UserControl_Show()
mIsVisible = True
End Sub

Private Sub UserControl_Terminate()
Debug.Print "DOMDisplay control terminated"
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub DOMGrid_Click()
DOMGrid.row = 1
DOMGrid.col = 0
End Sub

'@================================================================================
' MarketDepthListener Interface Members
'@================================================================================

Private Sub MarketDepthListener_resetMarketDepth( _
                ev As MarketDepthEvent)
reset
End Sub

Private Sub MarketDepthListener_setMarketDepthCell( _
                ev As MarketDepthEvent)
If mInitialPrice = 0 Then
    mInitialPrice = ev.price
    setupRows
    centreRow mInitialPrice
End If

setDOMCell ev.Type, ev.price, ev.size
End Sub

'@================================================================================
' mCentreTimer Event Handlers
'@================================================================================

Private Sub mCentreTimer_TimerExpired()
If mInitialPrice <> 0 Then centreRow mInitialPrice
Set mCentreTimer = Nothing
End Sub

'@================================================================================
' mTicker Event Handlers
'@================================================================================

Private Sub mTicker_Notification(ev As NotificationEvent)
If ev.eventCode = ApiNotifyCodes.ApiNotifyMarketDepthNotAvailable Then
    If Not mCentreTimer Is Nothing Then mCentreTimer.StopTimer
    Set mCentreTimer = Nothing
    finish
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

Private Property Let initialPrice(ByVal value As Double)
If mInitialPrice <> 0 Then Exit Property
mInitialPrice = value
End Property

Public Property Let numberOfRows(ByVal value As Long)

If value < 5 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "TradeBuildUI.DOMDisplay::numberOfRows()", _
                "Value must be >= 5"
End If

DOMGrid.Rows = value
End Property

Public Property Let ticker(ByVal value As ticker)
Set mTicker = value
Set mContract = mTicker.Contract
mPriceIncrement = mContract.ticksize

If mTicker.TradePrice <> 0 Then
    initialPrice = mTicker.TradePrice
ElseIf mTicker.BidPrice <> 0 Then
    initialPrice = mTicker.BidPrice
ElseIf mTicker.AskPrice <> 0 Then
    initialPrice = mTicker.AskPrice
ElseIf mTicker.closePrice <> 0 Then
    initialPrice = mTicker.AskPrice
End If

If mInitialPrice <> 0 Then
    setupRows
    
    If mTicker.TradePrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateLast, mTicker.TradePrice, mTicker.TradeSize
    End If
    If mTicker.AskPrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateAsk, mTicker.AskPrice, mTicker.AskSize
    End If
    If mTicker.BidPrice <> 0 Then
        setDOMCell DOMUpdateTypes.DOMUpdateBid, mTicker.BidPrice, mTicker.bidSize
    End If
    
    ' set off a timer before centring the display - otherwise it centres
    ' before the first resize
    Set mCentreTimer = CreateIntervalTimer(10)
    mCentreTimer.StartTimer
End If

mTicker.addMarketDepthListener Me

mTicker.RequestMarketDepth DOMEvents.DOMProcessedEvents, False

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub finish()
On Error GoTo Err
If Not mCentreTimer Is Nothing Then mCentreTimer.StopTimer
mTicker.removeMarketDepthListener Me
mTicker.CancelMarketDepth
Set mTicker = Nothing
Exit Sub
Err:
'ignore any errors
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


Private Function calcRowNumber(ByVal price As Double) As Long
calcRowNumber = ((mCeilingPrice - price) / mPriceIncrement) + 1
End Function

Private Sub centreRow(ByVal price As Double)
DOMGrid.TopRow = calcRowNumber(price) - Int(mNumberOfVisibleRows / 2)
End Sub

Private Sub checkEnoughRows(ByVal price As Double)
Dim i As Long
Dim rowprice As Double

If price = 0 Then Exit Sub

If (price - mBasePrice) / mPriceIncrement <= 5 Then
    ' add some new list entries at the start
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mBasePrice - (i * mPriceIncrement)
            DOMGrid.addItem ""
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, mTicker.formatPrice(rowprice)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, mTicker.formatPrice(rowprice)
        Next
        mBasePrice = rowprice
    Loop Until (price - mBasePrice) / mPriceIncrement > 5
    
    centreRow IIf(mCurrentLast <> 0, mCurrentLast, (mCeilingPrice + mBasePrice) / 2)
    DOMGrid.Redraw = True
End If

If (mCeilingPrice - price) / mPriceIncrement <= 5 Then
    ' add some new list entries at the end
    DOMGrid.Redraw = False
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mCeilingPrice + (i * mPriceIncrement)
            DOMGrid.addItem "", 1
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, mTicker.formatPrice(rowprice)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, mTicker.formatPrice(rowprice)
        Next
        mCeilingPrice = rowprice
    Loop Until (mCeilingPrice - price) / mPriceIncrement > 5

    centreRow IIf(mCurrentLast <> 0, mCurrentLast, (mCeilingPrice + mBasePrice) / 2)
    DOMGrid.Redraw = True
End If

End Sub

Private Sub clearDisplay(ByVal updateType As DOMUpdateTypes, ByVal price As Double)
checkEnoughRows price
Select Case updateType
Case DOMUpdateTypes.DOMUpdateAsk
    setCellContents calcRowNumber(price), DOMColumns.AskSize, ""
Case DOMUpdateTypes.DOMUpdateBid
    setCellContents calcRowNumber(price), DOMColumns.bidSize, ""
Case DOMUpdateTypes.DOMUpdateLast
    setCellContents calcRowNumber(price), DOMColumns.LastSize, ""
End Select
End Sub

Private Sub reset()
mHalted = True
DOMGrid.clear
ReDim mAskPrices(20) As Double
ReDim mBidPrices(20) As Double

mInitialPrice = mCurrentLast

mMaxAskPricesIndex = 0
mMaxBidPricesIndex = 0

mCurrentLast = 0#

setupRows
RaiseEvent Halted
End Sub

Private Sub resize()
Dim i As Long
Dim colWidth As Long

DOMGrid.Left = 0
DOMGrid.Top = 0
DOMGrid.Width = UserControl.Width
DOMGrid.Height = UserControl.Height

mNumberOfVisibleRows = Int(DOMGrid.Height / DOMGrid.rowHeight(1)) - 1

colWidth = (UserControl.ScaleWidth - ScrollbarWidth) / DOMGrid.Cols
For i = 0 To DOMGrid.Cols - 2
    DOMGrid.colWidth(i) = colWidth
Next
DOMGrid.colWidth(DOMGrid.Cols - 1) = UserControl.ScaleWidth - ScrollbarWidth - (DOMGrid.Cols - 1) * colWidth
End Sub

Private Sub setCellContents(ByVal row As Long, ByVal col As Long, ByVal value As String)
Dim currVal As String

DOMGrid.row = row
DOMGrid.col = col

currVal = DOMGrid.Text

If (currVal <> "" And value = "") Or _
    (currVal = "" And value <> "") _
Then
    If row <> 0 Then
        If value = "" Then
            DOMGrid.CellBackColor = GridColours.BGDefault
        Else
            Select Case col
            Case DOMColumns.bidSize
                DOMGrid.CellBackColor = GridColours.BGBid
            Case DOMColumns.LastSize
                DOMGrid.CellBackColor = GridColours.BGLast
                DOMGrid.CellFontBold = True
            Case DOMColumns.AskSize
                DOMGrid.CellBackColor = GridColours.BGAsk
            Case Else
                DOMGrid.CellBackColor = GridColours.BGDefault
            End Select
        End If
    End If
End If
DOMGrid.Text = value
End Sub

Private Sub setDOMCell( _
                ByVal updateType As DOMUpdateTypes, _
                ByVal price As Double, _
                ByVal size As Long)
If mHalted Then
    mHalted = False
    RaiseEvent Resumed
End If
If size > 0 Then
    setDisplay updateType, price, size
Else
    clearDisplay updateType, price
End If
End Sub
                
Private Sub setDisplay(ByVal updateType As DOMUpdateTypes, ByVal price As Double, ByVal size As Long)
checkEnoughRows price
Select Case updateType
Case DOMUpdateTypes.DOMUpdateAsk
    setCellContents calcRowNumber(price), DOMColumns.AskSize, size
Case DOMUpdateTypes.DOMUpdateBid
    setCellContents calcRowNumber(price), DOMColumns.bidSize, size
Case DOMUpdateTypes.DOMUpdateLast
    setCellContents calcRowNumber(price), DOMColumns.LastSize, size
    mCurrentLast = price
End Select
End Sub

Private Sub setupRows()
Dim i As Long
Dim price As Double

mBasePrice = mInitialPrice - (mPriceIncrement * Int(DOMGrid.Rows / 2))
mCeilingPrice = mBasePrice + (DOMGrid.Rows - 2) * mPriceIncrement

DOMGrid.Redraw = False

For i = DOMGrid.Rows - 1 To 1 Step -1
    price = mBasePrice + (DOMGrid.Rows - 1 - i) * mPriceIncrement
    setCellContents i, DOMColumns.PriceLeft, mTicker.formatPrice(price)
    setCellContents i, DOMColumns.PriceRight, mTicker.formatPrice(price)
Next

DOMGrid.Redraw = True

End Sub



