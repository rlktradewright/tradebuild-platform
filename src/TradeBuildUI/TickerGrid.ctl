VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.UserControl TickerGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin MSFlexGridLib.MSFlexGrid TickerGrid 
      Height          =   1335
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2355
      _Version        =   393216
   End
End
Attribute VB_Name = "TickerGrid"
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
' Interfaces
'================================================================================

Implements QuoteListener
Implements PriceChangeListener

'================================================================================
' Events
'================================================================================

Event Click()

'================================================================================
' Constants
'================================================================================

Private Const CellBackColorOdd As Long = &HF8F8F8
Private Const CellBackColorEven As Long = &HEEEEEE

Private Const GridRowsInitial As Long = 25
Private Const GridRowsIncrement As Long = 25

Private Const KeyDownShift As Integer = &H1
Private Const KeyDownCtrl As Integer = &H2
Private Const KeyDownAlt As Integer = &H4

Private Const TickerTableEntriesInitial As Long = 30
Private Const TickerTableEntriesIncrement As Long = 30

'================================================================================
' Enums
'================================================================================

Private Enum TickerGridColumns
    Selector
    TickerName
    currencyCode
    bidSize
    bid
    ask
    AskSize
    trade
    TradeSize
    volume
    Change
    ChangePercent
    highPrice
    lowPrice
    closePrice
    openInterest
    Description
    symbol
    sectype
    expiry
    exchange
    OptionRight
    strike
End Enum

' Character widths of the TickerGrid columns
Private Enum TickerGridColumnWidths
    SelectorWidth = 3
    NameWidth = 11
    CurrencyWidth = 5
    BidSizeWidth = 7
    BidWidth = 9
    AskWidth = 9
    AskSizeWidth = 7
    TradeWidth = 9
    TradeSIzeWidth = 7
    VolumeWidth = 9
    ChangeWidth = 7
    ChangePercentWidth = 7
    highWidth = 9
    LowWidth = 9
    CloseWidth = 9
    openInterestWidth = 9
    DescriptionWidth = 20
    SymbolWidth = 5
    SecTypeWidth = 10
    ExpiryWidth = 10
    ExchangeWidth = 10
    OptionRightWidth = 5
    StrikeWidth = 9
End Enum

Private Enum TickerGridSummaryColumns
    Selector
    TickerName
    bidSize
    bid
    ask
    AskSize
    trade
    TradeSize
    volume
    Change
    ChangePercent
    openInterest
End Enum

' Character widths of the TickerGrid columns (summary mode)
Private Enum TickerGridSummaryColumnWidths
    SelectorWidth = 3
    NameWidth = 13
    BidSizeWidth = 6
    BidWidth = 8
    AskWidth = 8
    AskSizeWidth = 6
    TradeWidth = 8
    TradeSIzeWidth = 6
    VolumeWidth = 8
    ChangeWidth = 8
    ChangePercentWidth = 8
    openInterestWidth
End Enum

'================================================================================
' Types
'================================================================================

Private Type TickerTableEntry
    theTicker               As ticker
    tickerGridRow           As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private WithEvents mTickers As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private mTickerTable() As TickerTableEntry

Private mLetterWidth As Single
Private mDigitWidth As Single

Private mNextGridRowIndex As Long

'/**
'  Contains an entry for each row in the grid. Set to 1 when the corresponding
'  grid row is selected
'*/
Private mSelectedRowsTable() As Long

Private mControlDown As Boolean
Private mShiftDown As Boolean
Private mAltDown As Boolean

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Dim widthString As String

ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
mNextGridRowIndex = 1

ReDim mSelectedRowsTable(GridRowsInitial - 1) As Long

widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = UserControl.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = UserControl.TextWidth(widthString) / Len(widthString)

setupDefaultTickerGrid

End Sub

Private Sub UserControl_Resize()
TickerGrid.Left = 0
TickerGrid.Top = 0
TickerGrid.Width = UserControl.Width
TickerGrid.Height = UserControl.Height
End Sub

'================================================================================
' PriceChangeListener Interface Members
'================================================================================

Private Sub PriceChangeListener_Change(ev As PriceChangeEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.Change
TickerGrid.Text = ev.ChangeString
If ev.Change >= 0 Then
    TickerGrid.CellBackColor = PositiveChangeBackColor
Else
    TickerGrid.CellBackColor = NegativeChangebackColor
End If
TickerGrid.CellForeColor = vbWhite

TickerGrid.col = TickerGridColumns.ChangePercent
TickerGrid.Text = Format(ev.ChangePercent, "0.0")
If ev.ChangePercent >= 0 Then
    TickerGrid.CellBackColor = PositiveChangeBackColor
Else
    TickerGrid.CellBackColor = NegativeChangebackColor
End If
TickerGrid.CellForeColor = vbWhite
End Sub

'================================================================================
' QuoteListener Interface Members
'================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.ask
TickerGrid.Text = ev.priceString
If ev.priceChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.priceChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If

TickerGrid.col = TickerGridColumns.AskSize
TickerGrid.Text = ev.size
If ev.sizeChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.sizeChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If
End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.bid
If ev.priceChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.priceChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If

TickerGrid.Text = ev.priceString
TickerGrid.col = TickerGridColumns.bidSize
TickerGrid.Text = ev.size
If ev.sizeChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.sizeChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If
End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.highPrice
TickerGrid.Text = ev.priceString
End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.lowPrice
TickerGrid.Text = ev.priceString
End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.openInterest
TickerGrid.Text = ev.size
End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.closePrice
TickerGrid.Text = ev.priceString
End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.trade
TickerGrid.Text = ev.priceString
If ev.priceChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.priceChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If

TickerGrid.col = TickerGridColumns.TradeSize
TickerGrid.Text = ev.size
If ev.sizeChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.sizeChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If
End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(lTicker.handle).tickerGridRow
TickerGrid.col = TickerGridColumns.volume
TickerGrid.Text = ev.size
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub TickerGrid_Click()
Dim row As Long
Dim rowSel As Long
Dim col As Long
Dim colSel As Long
Dim i As Long
row = TickerGrid.row
rowSel = TickerGrid.rowSel
col = TickerGrid.col
colSel = TickerGrid.colSel

If col = 1 And colSel = TickerGrid.Cols - 1 Then
    ' the user has clicked in the selector column
    If row = 1 And rowSel = TickerGrid.Rows - 1 Then
        ' the user has clicked on the top left cell so select all rows
        ' regardless of whether ctrl is down
        deselectSelectedRows
        
        TickerGrid.Redraw = False
        For i = 1 To mNextGridRowIndex - 1
            mSelectedRowsTable(i) = 1
            invertEntryColors i
        Next
        TickerGrid.Redraw = True
    Else
        If Not mControlDown Then
            deselectSelectedRows
            If row < mNextGridRowIndex Then
                mSelectedRowsTable(row) = 1
                invertEntryColors row
            End If
        Else
            If row < mNextGridRowIndex Then
                mSelectedRowsTable(row) = mSelectedRowsTable(row) Xor 1 ' toggle the entry
                invertEntryColors row
            End If
        End If
    End If
Else
    deselectSelectedRows
End If

RaiseEvent Click
End Sub

Private Sub TickerGrid_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
End Sub

Private Sub TickerGrid_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                x As Single, _
                y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
End Sub

'================================================================================
' mTickers Event Handlers
'================================================================================

Private Sub mTickers_StateChange(ev As TWUtilities.StateChangeEvent)
Dim lTicker As ticker
Dim handle As Long
Dim lContract As Contract
Dim gridRowIndex As Long
    

Set lTicker = ev.Source
handle = lTicker.handle
    

Select Case ev.State
Case TickerStateCreated

Case TickerStateStarting

Case TickerStateRunning
    
    If handle > UBound(mTickerTable) Then
        ReDim Preserve mTickerTable(UBound(mTickerTable) + TickerTableEntriesIncrement) As TickerTableEntry
    End If
    
    Set mTickerTable(handle).theTicker = lTicker
    
    If mNextGridRowIndex > TickerGrid.Rows - 5 Then
        TickerGrid.Rows = TickerGrid.Rows + GridRowsIncrement
        ReDim Preserve mSelectedRowsTable(UBound(mSelectedRowsTable) + GridRowsIncrement) As Long
    End If
    
    mTickerTable(handle).tickerGridRow = mNextGridRowIndex
    mNextGridRowIndex = mNextGridRowIndex + 1
    lTicker.addQuoteListener Me
    lTicker.addPriceChangeListener Me
    
    Set lContract = lTicker.Contract
    
    TickerGrid.row = mTickerTable(handle).tickerGridRow
    TickerGrid.rowdata(mTickerTable(handle).tickerGridRow) = handle
    
    TickerGrid.col = TickerGridColumns.currencyCode
    TickerGrid.Text = lContract.specifier.currencyCode
    
    TickerGrid.col = TickerGridColumns.Description
    TickerGrid.Text = lContract.Description
    
    TickerGrid.col = TickerGridColumns.exchange
    TickerGrid.Text = lContract.specifier.exchange
    
    TickerGrid.col = TickerGridColumns.expiry
    TickerGrid.Text = lContract.ExpiryDate
    
    TickerGrid.col = TickerGridColumns.OptionRight
    TickerGrid.Text = OptionRightToString(lContract.specifier.Right)
    
    TickerGrid.col = TickerGridColumns.sectype
    TickerGrid.Text = SecTypeToString(lContract.specifier.sectype)
    
    TickerGrid.col = TickerGridColumns.strike
    TickerGrid.Text = lContract.specifier.strike
    
    TickerGrid.col = TickerGridColumns.symbol
    TickerGrid.Text = lContract.specifier.symbol
    
    TickerGrid.col = TickerGridColumns.TickerName
    TickerGrid.Text = lContract.specifier.localSymbol
    
Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    ' if the ticker was stopped by the application via a call to Ticker.topTicker (rather
    ' tha via this control), the entry will still be in the grid so remove it
    If Not mTickerTable(handle).theTicker Is Nothing Then
        gridRowIndex = mTickerTable(handle).tickerGridRow
        removeTicker handle
        setGridRowBackColors gridRowIndex
    End If
End Select
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get SelectedTickers() As SelectedTickers
Dim i As Long

Set SelectedTickers = New SelectedTickers

For i = 1 To mNextGridRowIndex - 1
    If mSelectedRowsTable(i) = 1 Then
        SelectedTickers.add mTickerTable(TickerGrid.rowdata(i)).theTicker
    End If
Next

End Property

'================================================================================
' Methods
'================================================================================

Public Sub finish()
On Error GoTo Err
StopAllTickers
Set mTickers = Nothing
ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
Exit Sub
Err:
'ignore any errors
End Sub

Public Sub monitorWorkspace( _
                ByVal pWorkspace As WorkSpace)
If Not mTickers Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                            "TradeBuildUI.TickerGrid::monitorWorkspace", _
                                            "A workspace is already being monitored"
Set mTickers = pWorkspace.Tickers
End Sub
                
Public Sub StopAllTickers()
Dim i As Long

TickerGrid.Redraw = False

' do this in reverse order - most efficient when all tickers are being stopped
For i = mNextGridRowIndex - 1 To 1 Step -1
    stopTicker i
Next
TickerGrid.Redraw = True

setGridRowBackColors 1
End Sub

Public Sub StopSelectedTickers()
Dim i As Long
Dim lowestIndex As Long

TickerGrid.Redraw = False

' do this in reverse order - most efficient when all tickers are being stopped
For i = mNextGridRowIndex - 1 To 1 Step -1
    If mSelectedRowsTable(i) = 1 Then
        stopTicker i
        lowestIndex = i
    End If
Next
TickerGrid.Redraw = True

setGridRowBackColors lowestIndex
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub deselectSelectedRows()
Dim i As Long
TickerGrid.Redraw = False
For i = 0 To mNextGridRowIndex - 1
    If mSelectedRowsTable(i) <> 0 Then
        invertEntryColors i
        mSelectedRowsTable(i) = 0
    End If
Next
TickerGrid.Redraw = True
End Sub

'/**
'   Inverts the foreground and background colors for the current grid cell
'*/
Private Sub invertCellColors()
foreColor = IIf(TickerGrid.CellForeColor = 0, TickerGrid.foreColor, TickerGrid.CellForeColor)
If foreColor = SystemColorConstants.vbWindowText Then
    TickerGrid.CellForeColor = SystemColorConstants.vbHighlightText
ElseIf foreColor = SystemColorConstants.vbHighlightText Then
    TickerGrid.CellForeColor = SystemColorConstants.vbWindowText
ElseIf foreColor > 0 Then
    TickerGrid.CellForeColor = IIf((foreColor Xor &HFFFFFF) = 0, 1, foreColor Xor &HFFFFFF)
End If

backColor = IIf(TickerGrid.CellBackColor = 0, TickerGrid.backColor, TickerGrid.CellBackColor)
If backColor = SystemColorConstants.vbWindowBackground Then
    TickerGrid.CellBackColor = SystemColorConstants.vbHighlight
ElseIf backColor = SystemColorConstants.vbHighlight Then
    TickerGrid.CellBackColor = SystemColorConstants.vbWindowBackground
ElseIf backColor > 0 Then
    TickerGrid.CellBackColor = IIf((backColor Xor &HFFFFFF) = 0, 1, backColor Xor &HFFFFFF)
End If
End Sub

Private Sub invertEntryColors(ByVal rowNumber As Long)
Dim i As Long

If rowNumber < 0 Then Exit Sub

TickerGrid.row = rowNumber

For i = 1 To TickerGrid.Cols - 1
    TickerGrid.col = i
    If TickerGrid.CellFontBold Then
        TickerGrid.CellFontBold = False
    Else
        TickerGrid.CellFontBold = True
    End If
Next

TickerGrid.col = TickerGridColumns.TickerName
invertCellColors

TickerGrid.col = TickerGridColumns.currencyCode
invertCellColors

TickerGrid.col = TickerGridColumns.Description
invertCellColors

TickerGrid.col = TickerGridColumns.exchange
invertCellColors

TickerGrid.col = TickerGridColumns.sectype
invertCellColors

TickerGrid.col = TickerGridColumns.symbol
invertCellColors

End Sub

Private Sub removeTicker( _
                ByVal handle As Long)
Dim gridRowIndex As Long
Dim i As Long
Dim rowdata As Long

gridRowIndex = mTickerTable(handle).tickerGridRow
TickerGrid.RemoveItem gridRowIndex
Set mTickerTable(handle).theTicker = Nothing
mTickerTable(handle).tickerGridRow = 0

mSelectedRowsTable(gridRowIndex) = 0

For i = gridRowIndex + 1 To mNextGridRowIndex - 1
    mSelectedRowsTable(i - 1) = mSelectedRowsTable(i)
Next

mNextGridRowIndex = mNextGridRowIndex - 1

For i = gridRowIndex To mNextGridRowIndex - 1
    rowdata = TickerGrid.rowdata(i)
    mTickerTable(rowdata).tickerGridRow = i
Next
End Sub

Private Sub setGridRowBackColors( _
                ByVal startingIndex As Long)
Dim i As Long
Dim lTicker As ticker

TickerGrid.Redraw = False

For i = startingIndex To TickerGrid.Rows - 1
    TickerGrid.row = i
    TickerGrid.col = 1
    TickerGrid.rowSel = i
    TickerGrid.colSel = TickerGrid.Cols - 1
    TickerGrid.CellBackColor = IIf(i Mod 2 = 0, CellBackColorEven, CellBackColorOdd)
    
    If i <= mNextGridRowIndex - 1 Then
        Set lTicker = mTickerTable(TickerGrid.rowdata(i)).theTicker
        
        If lTicker.Change > 0 Then
            TickerGrid.col = TickerGridColumns.Change
            TickerGrid.colSel = TickerGridColumns.Change
            TickerGrid.CellBackColor = PositiveChangeBackColor
            
            TickerGrid.col = TickerGridColumns.ChangePercent
            TickerGrid.colSel = TickerGridColumns.ChangePercent
            TickerGrid.CellBackColor = PositiveChangeBackColor
        Else
            TickerGrid.col = TickerGridColumns.Change
            TickerGrid.colSel = TickerGridColumns.Change
            TickerGrid.CellBackColor = NegativeChangebackColor
            
            TickerGrid.col = TickerGridColumns.ChangePercent
            TickerGrid.colSel = TickerGridColumns.ChangePercent
            TickerGrid.CellBackColor = NegativeChangebackColor
        End If
        
        TickerGrid.CellForeColor = vbWhite
    End If
Next

TickerGrid.Redraw = True
End Sub

Private Sub setupDefaultTickerGrid()

With TickerGrid
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .HighLight = flexHighlightNever
    
    .Cols = 2
    .Rows = GridRowsInitial
    .FixedRows = 1
    .FixedCols = 1
End With
    
setupTickerGridColumn 0, TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, "", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, "Name", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.currencyCode, TickerGridColumnWidths.CurrencyWidth, "Curr", True, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.bidSize, TickerGridColumnWidths.BidSizeWidth, "Bid size", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.bid, TickerGridColumnWidths.BidWidth, "Bid", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.ask, TickerGridColumnWidths.AskWidth, "Ask", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, "Ask size", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.trade, TickerGridColumnWidths.TradeWidth, "Last", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSIzeWidth, "Last size", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.volume, TickerGridColumnWidths.VolumeWidth, "Volume", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, "Chg", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, "Chg %", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.highPrice, TickerGridColumnWidths.highWidth, "High", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.lowPrice, TickerGridColumnWidths.LowWidth, "Low", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.closePrice, TickerGridColumnWidths.CloseWidth, "Close", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.openInterest, TickerGridColumnWidths.openInterestWidth, "Open interest", False, AlignmentSettings.flexAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, "Description", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.symbol, TickerGridColumnWidths.SymbolWidth, "Symbol", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.sectype, TickerGridColumnWidths.SecTypeWidth, "Sec Type", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.expiry, TickerGridColumnWidths.ExpiryWidth, "Expiry", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.exchange, TickerGridColumnWidths.ExchangeWidth, "Exchange", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, "Right", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.strike, TickerGridColumnWidths.StrikeWidth, "Strike", False, AlignmentSettings.flexAlignLeftCenter

setGridRowBackColors 1
End Sub

Private Sub setupSummaryTickerGrid()
With TickerGrid
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .HighLight = flexHighlightNever
    
    .Cols = 2
    .Rows = GridRowsInitial
    .FixedRows = 1
    .FixedCols = 1
End With
    
setupTickerGridColumn 0, TickerGridSummaryColumns.Selector, TickerGridSummaryColumnWidths.SelectorWidth, "", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.TickerName, TickerGridSummaryColumnWidths.NameWidth, "Name", True, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.bidSize, TickerGridSummaryColumnWidths.BidSizeWidth, "Bid size", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.bid, TickerGridSummaryColumnWidths.BidWidth, "Bid", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.ask, TickerGridSummaryColumnWidths.AskWidth, "Ask", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.AskSize, TickerGridSummaryColumnWidths.AskSizeWidth, "Ask size", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.trade, TickerGridSummaryColumnWidths.TradeWidth, "Last", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.TradeSize, TickerGridSummaryColumnWidths.TradeSIzeWidth, "Last size", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.volume, TickerGridSummaryColumnWidths.VolumeWidth, "Volume", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.Change, TickerGridSummaryColumnWidths.ChangeWidth, "Change", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.ChangePercent, TickerGridSummaryColumnWidths.ChangePercentWidth, "Change %", False, AlignmentSettings.flexAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.openInterest, TickerGridSummaryColumnWidths.openInterestWidth, "Open interest", False, AlignmentSettings.flexAlignLeftCenter

setGridRowBackColors 1
End Sub

Private Sub setupTickerGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As AlignmentSettings)
    
Dim lColumnWidth As Long

With TickerGrid
    .row = rowNumber
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .colWidth(columnNumber) = 0
    End If
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
'    If .ColWidth(columnNumber) < lColumnWidth Then
        .colWidth(columnNumber) = lColumnWidth
'    End If
    
    .ColAlignment(columnNumber) = align
    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With
End Sub
                
Private Sub stopTicker( _
                ByVal gridRowIndex As Long)
Dim lTicker As ticker
Dim handle As Long

If mSelectedRowsTable(gridRowIndex) = 0 Then Exit Sub

Set lTicker = mTickerTable(TickerGrid.rowdata(gridRowIndex)).theTicker
handle = lTicker.handle
lTicker.removeQuoteListener Me
lTicker.removePriceChangeListener Me

removeTicker handle

lTicker.stopTicker
End Sub



