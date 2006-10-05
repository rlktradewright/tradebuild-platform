VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.UserControl DOMDisplay 
   ClientHeight    =   1725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3795
   ScaleHeight     =   1725
   ScaleWidth      =   3795
   Begin MSFlexGridLib.MSFlexGrid DOMGrid 
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   2143
      _Version        =   393216
      ScrollBars      =   2
   End
End
Attribute VB_Name = "DOMDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Implements TradeBuild.ProcessedMarketDepthListener

'================================================================================
' Events
'================================================================================

Event Halted()
Event Resumed()

'================================================================================
' Constants
'================================================================================

Private Const ScrollbarWidth As Long = 370  ' value discovered by trial and error!

'================================================================================
' Enums
'================================================================================

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

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTicker As ticker
Attribute mTicker.VB_VarHelpID = -1

Private mContract As Contract
Private mInitialPrice As Double
Private mPriceIncrement As Double
Private mNumberOfRows As Long
Private mNumberOfVisibleRows As Long
Private mBasePrice As Double
Private mCeilingPrice As Double

Private mAskPrices() As Double
Private mMaxAskPricesIndex As Long
Private mBidPrices() As Double
Private mMaxBidPricesIndex As Long

Private mCurrentLast As Double

Private mHalted As Boolean

Private WithEvents mTimer As TimerUtils.IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub UserControl_Initialize()

ReDim mAskPrices(20) As Double
ReDim mBidPrices(20) As Double

DOMGrid.AllowUserResizing = flexResizeColumns
DOMGrid.BackColorFixed = vbButtonFace
DOMGrid.BackColorSel = GridColours.BGDefault
DOMGrid.backColor = GridColours.BGDefault
DOMGrid.ForeColorSel = DOMGrid.foreColor
DOMGrid.Rows = 2
DOMGrid.Cols = 5
DOMGrid.FixedCols = 0
DOMGrid.FixedRows = 1
DOMGrid.FocusRect = flexFocusNone
DOMGrid.ScrollBars = flexScrollBarVertical



setCellContents 0, DOMColumns.PriceLeft, "Price"
DOMGrid.colWidth(DOMColumns.PriceLeft) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceLeft) = flexAlignRightCenter

setCellContents 0, DOMColumns.bidSize, "Bids"
DOMGrid.colWidth(DOMColumns.bidSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.bidSize) = flexAlignCenterCenter

setCellContents 0, DOMColumns.LastSize, "Last"
DOMGrid.colWidth(DOMColumns.LastSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.LastSize) = flexAlignCenterCenter

setCellContents 0, DOMColumns.AskSize, "Asks"
DOMGrid.colWidth(DOMColumns.AskSize) = 14 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.AskSize) = flexAlignCenterCenter

setCellContents 0, DOMColumns.PriceRight, "Price"
DOMGrid.colWidth(DOMColumns.PriceRight) = 24 * DOMGrid.Width / 100
DOMGrid.ColAlignment(DOMColumns.PriceRight) = flexAlignLeftCenter

mNumberOfRows = 50
End Sub

Private Sub UserControl_Resize()
Dim i As Long
Dim colWidth As Long

DOMGrid.Left = 0
DOMGrid.Top = 0
DOMGrid.Width = UserControl.Width
DOMGrid.Height = UserControl.Height

mNumberOfVisibleRows = Int(DOMGrid.Height / DOMGrid.RowHeight(1))

colWidth = (UserControl.ScaleWidth - ScrollbarWidth) / DOMGrid.Cols
For i = 0 To DOMGrid.Cols - 2
    DOMGrid.colWidth(i) = colWidth
Next
DOMGrid.colWidth(DOMGrid.Cols - 1) = UserControl.ScaleWidth - ScrollbarWidth - (DOMGrid.Cols - 1) * colWidth
End Sub

Private Sub UserControl_Terminate()
Debug.Print "DOMDisplay control terminated"
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub DOMGrid_Click()
DOMGrid.row = 1
DOMGrid.col = 0
End Sub

'================================================================================
' ProcessedMarketDepthListener Interface Members
'================================================================================

Private Sub ProcessedMarketDepthListener_resetMarketDepth( _
                ev As TradeBuild.ProcessedMarketDepthEvent)
reset
End Sub

Private Sub ProcessedMarketDepthListener_setMarketDepthCell( _
                ev As TradeBuild.ProcessedMarketDepthEvent)
setDOMCell ev.side, ev.price, ev.size
End Sub

'================================================================================
' mTicker Event Handlers
'================================================================================

Private Sub mTicker_Error(ev As TradeBuild.ErrorEvent)
If ev.errorCode = ApiErrorCodes.ApiErrMarketDepthNotAvailable Then
    If Not mTimer Is Nothing Then mTimer.StopTimer
    finish
End If
End Sub

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Set mTimer = Nothing
centreRow mInitialPrice
End Sub

'================================================================================
' Properties
'================================================================================

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

mNumberOfRows = value
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
End If

setupRows

If mTicker.TradePrice <> 0 Then
    setDOMCell DOMSides.DOMLast, mTicker.TradePrice, mTicker.TradeSize
End If
If mTicker.AskPrice <> 0 Then
    setDOMCell DOMSides.DOMAsk, mTicker.AskPrice, mTicker.AskSize
End If
If mTicker.BidPrice <> 0 Then
    setDOMCell DOMSides.DOMBid, mTicker.BidPrice, mTicker.bidSize
End If

mTicker.addProcessedMarketDepthListener Me

mTicker.RequestMarketDepth DOMEvents.DOMProcessedEvents, False

' set off a timer before centring the display - otherwise it centres
' before the first resize
Set mTimer = New TimerUtils.IntervalTimer
mTimer.TimerIntervalMillisecs = 10
mTimer.RepeatNotifications = False
mTimer.StartTimer
End Property

'================================================================================
' Methods
'================================================================================

Public Sub finish()
On Error GoTo Err
mTicker.removeProcessedMarketDepthListener Me
mTicker.CancelMarketDepth
Set mTicker = Nothing
Exit Sub
Err:
'ignore any errors
End Sub

'================================================================================
' Helper Functions
'================================================================================


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
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mBasePrice - (i * mPriceIncrement)
            DOMGrid.AddItem ""
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, mTicker.formatPrice(rowprice)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, mTicker.formatPrice(rowprice)
        Next
        mBasePrice = rowprice
    Loop Until (price - mBasePrice) / mPriceIncrement > 5
    
    DOMGrid.col = DOMColumns.PriceLeft
    DOMGrid.Sort = flexSortNumericDescending
    centreRow IIf(mCurrentLast <> 0, mCurrentLast, (mCeilingPrice + mBasePrice) / 2)
    
End If

If (mCeilingPrice - price) / mPriceIncrement <= 5 Then
    ' add some new list entries at the end
    Do
        For i = 1 To Int(mNumberOfVisibleRows / 2)
            rowprice = mCeilingPrice + (i * mPriceIncrement)
            DOMGrid.AddItem ""
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceLeft, mTicker.formatPrice(rowprice)
            setCellContents DOMGrid.Rows - 1, DOMColumns.PriceRight, mTicker.formatPrice(rowprice)
        Next
        mCeilingPrice = rowprice
    Loop Until (mCeilingPrice - price) / mPriceIncrement > 5

    DOMGrid.col = DOMColumns.PriceLeft
    DOMGrid.Sort = flexSortNumericDescending
    centreRow IIf(mCurrentLast <> 0, mCurrentLast, (mCeilingPrice + mBasePrice) / 2)

End If

End Sub

Private Sub clearDisplay(ByVal side As DOMSides, ByVal price As Double)
checkEnoughRows price
Select Case side
Case DOMSides.DOMAsk
    setCellContents calcRowNumber(price), DOMColumns.AskSize, ""
Case DOMSides.DOMBid
    setCellContents calcRowNumber(price), DOMColumns.bidSize, ""
Case DOMSides.DOMLast
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

Private Sub setCellContents(ByVal row As Long, ByVal col As Long, ByVal value As String)
Dim currVal As String

currVal = DOMGrid.TextMatrix(row, col)

If (currVal <> "" And value = "") Or _
    (currVal = "" And value <> "") _
Then
    If row <> 0 Then
        DOMGrid.row = row
        DOMGrid.col = col
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
        DOMGrid.row = 1
        DOMGrid.col = 0
    End If
End If
DOMGrid.TextMatrix(row, col) = value
End Sub

Private Sub setDOMCell( _
                ByVal side As DOMSides, _
                ByVal price As Double, _
                ByVal size As Long)
If mHalted Then
    mHalted = False
    RaiseEvent Resumed
End If
If size > 0 Then
    setDisplay side, price, size
Else
    clearDisplay side, price
End If
End Sub
                
Private Sub setDisplay(ByVal side As DOMSides, ByVal price As Double, ByVal size As Long)
checkEnoughRows price
Select Case side
Case DOMSides.DOMAsk
    setCellContents calcRowNumber(price), DOMColumns.AskSize, size
Case DOMSides.DOMBid
    setCellContents calcRowNumber(price), DOMColumns.bidSize, size
Case DOMSides.DOMLast
    setCellContents calcRowNumber(price), DOMColumns.LastSize, size
    mCurrentLast = price
End Select
End Sub

Private Sub setupRows()
Dim i As Long
Dim price As Double
Dim currentGridHeight  As Long

mBasePrice = mInitialPrice - (mPriceIncrement * Int(mNumberOfRows / 2))

For i = 0 To mNumberOfRows - 1
    price = mBasePrice + (i * mPriceIncrement)
    DOMGrid.AddItem "", 1
    setCellContents 1, DOMColumns.PriceLeft, mTicker.formatPrice(price)
    setCellContents 1, DOMColumns.PriceRight, mTicker.formatPrice(price)
    
Next
mCeilingPrice = price

End Sub



