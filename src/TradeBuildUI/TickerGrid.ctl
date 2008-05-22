VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#27.6#0"; "TWControls10.ocx"
Begin VB.UserControl TickerGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin TWControls10.TWGrid TickerGrid 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   4683
   End
End
Attribute VB_Name = "TickerGrid"
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
' Interfaces
'@================================================================================

Implements QuoteListener
Implements PriceChangeListener

'@================================================================================
' Events
'@================================================================================

Event Click()

'@================================================================================
' Constants
'@================================================================================

Private Const CellBackColorOdd As Long = &HF8F8F8
Private Const CellBackColorEven As Long = &HEEEEEE

Private Const GridRowsInitial As Long = 25
Private Const GridRowsIncrement As Long = 25

Private Const KeyDownShift As Integer = &H1
Private Const KeyDownCtrl As Integer = &H2
Private Const KeyDownAlt As Integer = &H4

Private Const TickerTableEntriesInitial As Long = 4
Private Const TickerTableEntriesGrowthFactor As Long = 2

'@================================================================================
' Enums
'@================================================================================

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
    MaxColumn = strike
End Enum

' Character widths of the TickerGrid columns
Private Enum TickerGridColumnWidths
    SelectorWidth = 3
    NameWidth = 11
    CurrencyWidth = 5
    BidSizeWidth = 8
    BidWidth = 9
    AskWidth = 9
    AskSizeWidth = 8
    TradeWidth = 9
    TradeSizeWidth = 8
    VolumeWidth = 10
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
    MaxSummaryColumn = openInterest
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
    TradeSizeWidth = 6
    VolumeWidth = 8
    ChangeWidth = 8
    ChangePercentWidth = 8
    openInterestWidth
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type TickerTableEntry
    theTicker               As ticker
    tickerGridRow           As Long
    nextSelected            As Long
    prevSelected            As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTickers As Tickers
Attribute mTickers.VB_VarHelpID = -1
Private mTickerTable() As TickerTableEntry

Private mLetterWidth As Single
Private mDigitWidth As Single

Private mNextGridRowIndex As Long

Private mControlDown As Boolean
Private mShiftDown As Boolean
Private mAltDown As Boolean

Private mEventCount As Long

Private WithEvents mCountTimer As IntervalTimer
Attribute mCountTimer.VB_VarHelpID = -1

Private mLogger As Logger

Private mColumnMap() As Long

Private mFirstSelected As Long

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Dim widthString As String

ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
mNextGridRowIndex = 1

mTickerTable(0).nextSelected = -1
mFirstSelected = 0

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

'@================================================================================
' PriceChangeListener Interface Members
'@================================================================================

Private Sub PriceChangeListener_Change(ev As PriceChangeEvent)
Dim lTicker As ticker
Set lTicker = ev.Source

TickerGrid.row = mTickerTable(getTickerIndexFromHandle(lTicker.handle)).tickerGridRow
TickerGrid.col = mColumnMap(TickerGridColumns.Change)
TickerGrid.Text = ev.ChangeString
If ev.Change >= 0 Then
    TickerGrid.CellBackColor = PositiveChangeBackColor
Else
    TickerGrid.CellBackColor = NegativeChangebackColor
End If
TickerGrid.CellForeColor = vbWhite
incrementEventCount

TickerGrid.col = mColumnMap(TickerGridColumns.ChangePercent)
TickerGrid.Text = Format(ev.ChangePercent, "0.0")
If ev.ChangePercent >= 0 Then
    TickerGrid.CellBackColor = PositiveChangeBackColor
Else
    TickerGrid.CellBackColor = NegativeChangebackColor
End If
TickerGrid.CellForeColor = vbWhite

incrementEventCount
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.ask)
displaySize ev, mColumnMap(TickerGridColumns.AskSize)

End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.bid)
displaySize ev, mColumnMap(TickerGridColumns.bidSize)

End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.highPrice)

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.lowPrice)

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

displaySize ev, mColumnMap(TickerGridColumns.openInterest)

End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.closePrice)

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.trade)
displaySize ev, mColumnMap(TickerGridColumns.TradeSize)

End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)

displaySize ev, mColumnMap(TickerGridColumns.volume)

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TickerGrid_ColMoved( _
                ByVal fromCol As Long, _
                ByVal toCol As Long)
Dim i As Long

If fromCol < toCol Then
    For i = fromCol To toCol
        mColumnMap(TickerGrid.ColData(i)) = i
    Next
Else
    For i = toCol To fromCol
        mColumnMap(TickerGrid.ColData(i)) = i
    Next
End If

End Sub

Private Sub TickerGrid_Click()
Dim row As Long
Dim rowSel As Long
Dim col As Long
Dim colSel As Long
row = TickerGrid.row
rowSel = TickerGrid.rowSel
col = TickerGrid.col
colSel = TickerGrid.colSel

If col = 1 And colSel = TickerGrid.Cols - 1 Then
    ' the user has clicked in the selector column
    If row = 1 And rowSel = TickerGrid.Rows - 1 Then
        ' the user has clicked on the top left cell so select all rows
        ' regardless of whether ctrl is down
        deselectSelectedTickers
        
        selectAllTickers
        
    Else
        If Not mControlDown Then
            deselectSelectedTickers
            selectRow row
        Else
            toggleRowSelection row
        End If
    End If
Else
    deselectSelectedTickers
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

Private Sub TickerGrid_RowMoved( _
                ByVal fromRow As Long, _
                ByVal toRow As Long)
Dim i As Long

If fromRow < toRow Then
    For i = fromRow To toRow
        mTickerTable(TickerGrid.rowdata(i)).tickerGridRow = i
    Next
Else
    For i = toRow To fromRow
        mTickerTable(TickerGrid.rowdata(i)).tickerGridRow = i
    Next
End If

End Sub

Private Sub TickerGrid_RowMoving( _
                ByVal fromRow As Long, _
                ByVal toRow As Long, _
                Cancel As Boolean)
If toRow > mNextGridRowIndex Then Cancel = True
End Sub

'@================================================================================
' mCountTimer Event Handlers
'@================================================================================

Private Sub mCountTimer_TimerExpired()
mLogger.Log LogLevelMediumDetail, "TickerGrid: events per second=" & mEventCount / 10
mEventCount = 0
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_StateChange(ev As StateChangeEvent)
Dim lTicker As ticker
Dim index As Long
Dim lContract As Contract
    

Set lTicker = ev.Source
If lTicker.isHistorical Then Exit Sub

index = getTickerIndexFromHandle(lTicker.handle)
    
Select Case ev.state
Case TickerStateCreated

Case TickerStateStarting
    
    If index > UBound(mTickerTable) Then
        ReDim Preserve mTickerTable((UBound(mTickerTable) + 1) * TickerTableEntriesGrowthFactor - 1) As TickerTableEntry
    End If
    
    Set mTickerTable(index).theTicker = lTicker
    mTickerTable(index).nextSelected = -1
    mTickerTable(index).prevSelected = -1
    
    If mNextGridRowIndex > TickerGrid.Rows - 5 Then
        TickerGrid.Rows = TickerGrid.Rows + GridRowsIncrement
    End If
    
    mTickerTable(index).tickerGridRow = mNextGridRowIndex
    mNextGridRowIndex = mNextGridRowIndex + 1
    lTicker.addQuoteListener Me
    lTicker.addPriceChangeListener Me

    TickerGrid.row = mTickerTable(index).tickerGridRow
    TickerGrid.rowdata(mTickerTable(index).tickerGridRow) = index
    
    TickerGrid.col = mColumnMap(TickerGridColumns.TickerName)
    TickerGrid.Text = "Starting"
    
Case TickerStateRunning
    
    Set lContract = lTicker.Contract
    
    TickerGrid.row = mTickerTable(index).tickerGridRow
    
    TickerGrid.col = mColumnMap(TickerGridColumns.currencyCode)
    TickerGrid.Text = lContract.specifier.currencyCode
    
    TickerGrid.col = mColumnMap(TickerGridColumns.Description)
    TickerGrid.Text = lContract.Description
    
    TickerGrid.col = mColumnMap(TickerGridColumns.exchange)
    TickerGrid.Text = lContract.specifier.exchange
    
    TickerGrid.col = mColumnMap(TickerGridColumns.expiry)
    TickerGrid.Text = lContract.expiryDate
    
    TickerGrid.col = mColumnMap(TickerGridColumns.OptionRight)
    TickerGrid.Text = OptionRightToString(lContract.specifier.Right)
    
    TickerGrid.col = mColumnMap(TickerGridColumns.sectype)
    TickerGrid.Text = SecTypeToString(lContract.specifier.sectype)
    
    TickerGrid.col = mColumnMap(TickerGridColumns.strike)
    TickerGrid.Text = lContract.specifier.strike
    
    TickerGrid.col = mColumnMap(TickerGridColumns.symbol)
    TickerGrid.Text = lContract.specifier.symbol
    
    TickerGrid.col = mColumnMap(TickerGridColumns.TickerName)
    TickerGrid.Text = lContract.specifier.localSymbol
    
Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    ' if the ticker was stopped by the application via a call to Ticker.topTicker (rather
    ' tha via this control), the entry will still be in the grid so remove it
    If Not mTickerTable(index).theTicker Is Nothing Then
        removeTicker index
    End If
End Select
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get selectedTickers() As selectedTickers
Dim index As Long

Set selectedTickers = New selectedTickers

index = mFirstSelected

Do While index <> 0
    selectedTickers.add mTickerTable(index).theTicker
    index = mTickerTable(index).nextSelected
Loop

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub deselectSelectedTickers()
Do While mFirstSelected <> 0
    deselectTicker mFirstSelected
Loop
End Sub

Public Sub finish()
On Error GoTo Err
stopAllTickers
Set mTickers = Nothing
ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
If Not mCountTimer Is Nothing Then mCountTimer.StopTimer
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

Set mLogger = GetLogger("diag.tradebuild.tradebuildui")

Set mCountTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds, 10000)
mCountTimer.StartTimer

End Sub
                
Public Sub selectAllTickers()
Dim i As Long
For i = 1 To mNextGridRowIndex - 1
    selectRow i
Next
End Sub

Public Sub stopAllTickers()
Dim i As Long

TickerGrid.Redraw = False

' do this in reverse order - most efficient when all tickers are being stopped
For i = TickerGrid.Rows - 1 To 1 Step -1
    If TickerGrid.rowdata(i) <> 0 Then
        stopTicker TickerGrid.rowdata(i)
    End If
Next
TickerGrid.Redraw = True

End Sub

Public Sub stopSelectedTickers()

TickerGrid.Redraw = False

Do While mFirstSelected <> 0
    stopTicker mFirstSelected
Loop

TickerGrid.Redraw = True

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub deselectRow( _
                ByVal row As Long)
deselectTicker TickerGrid.rowdata(row)
End Sub

Private Sub deselectTicker( _
                ByVal index As Long)
If isTickerSelected(index) Then
    mTickerTable(mTickerTable(index).nextSelected).prevSelected = mTickerTable(index).prevSelected
    If mTickerTable(index).prevSelected = -1 Then
        mFirstSelected = mTickerTable(index).nextSelected
    Else
        mTickerTable(mTickerTable(index).prevSelected).nextSelected = mTickerTable(index).nextSelected
    End If
    mTickerTable(index).nextSelected = -1
    mTickerTable(index).prevSelected = -1
    highlightRow mTickerTable(index).tickerGridRow
End If
End Sub

Private Sub displayPrice( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(getTickerIndexFromHandle(lTicker.handle)).tickerGridRow
TickerGrid.col = col
TickerGrid.Text = ev.priceString
If ev.priceChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.priceChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If

incrementEventCount
End Sub

Private Sub displaySize( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = mTickerTable(getTickerIndexFromHandle(lTicker.handle)).tickerGridRow
TickerGrid.col = col
TickerGrid.Text = ev.size
If ev.sizeChange = ValueChangeUp Then
    TickerGrid.CellForeColor = IncreasedValueColor
ElseIf ev.sizeChange = ValueChangeDown Then
    TickerGrid.CellForeColor = DecreasedValueColor
End If

incrementEventCount
End Sub

Private Function getTickerIndexFromHandle( _
                ByVal handle As Long) As Long
' allow for the fact that the first tickertable entry is not used - it is the
' terminator of the selected entries chain
getTickerIndexFromHandle = handle + 1
End Function

Private Sub incrementEventCount()
mEventCount = mEventCount + 1
End Sub

Private Sub highlightRow(ByVal rowNumber As Long)
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

TickerGrid.col = mColumnMap(TickerGridColumns.TickerName)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.currencyCode)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.Description)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.exchange)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.sectype)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.symbol)
TickerGrid.InvertCellColors

End Sub

Private Function isRowSelected( _
                ByVal row As Long)
isRowSelected = isTickerSelected(TickerGrid.rowdata(row))
End Function

Private Function isTickerSelected( _
                ByVal index As Long)
If mTickerTable(index).nextSelected <> -1 Then isTickerSelected = True
End Function

Private Sub removeTicker( _
                ByVal index As Long)
Dim gridRowIndex As Long
Dim i As Long
Dim rowdata As Long

deselectTicker index

gridRowIndex = mTickerTable(index).tickerGridRow

TickerGrid.RemoveItem gridRowIndex
mNextGridRowIndex = mNextGridRowIndex - 1

Set mTickerTable(index).theTicker = Nothing
mTickerTable(index).tickerGridRow = 0

For i = gridRowIndex To mNextGridRowIndex - 1
    rowdata = TickerGrid.rowdata(i)
    mTickerTable(rowdata).tickerGridRow = i
Next

End Sub

Private Sub selectRow( _
                ByVal row As Long)
selectTicker TickerGrid.rowdata(row)
End Sub

Private Sub selectTicker( _
                ByVal index As Long)
If Not mTickerTable(index).theTicker Is Nothing Then
    mTickerTable(index).nextSelected = mFirstSelected
    mTickerTable(index).prevSelected = -1
    mTickerTable(mFirstSelected).prevSelected = index
    mFirstSelected = index
    highlightRow mTickerTable(index).tickerGridRow
End If
End Sub

Private Sub setupColumnMap( _
                    ByVal maxIndex As Long)
Dim i As Long
ReDim mColumnMap(maxIndex) As Long
For i = 0 To UBound(mColumnMap)
    mColumnMap(i) = i
Next
End Sub

Private Sub setupDefaultTickerGrid()

With TickerGrid
    .RowBackColorEven = CellBackColorEven
    .RowBackColorOdd = CellBackColorOdd
    .AllowUserReordering = TwGridReorderBoth
    .AllowBigSelection = True
    .AllowUserResizing = TwGridResizeBoth
    .RowSizingMode = TwGridRowSizeAll
    .FillStyle = TwGridFillRepeat
    .FocusRect = TwGridFocusNone
    .HighLight = TwGridHighlightNever
    
    .Cols = 2
    .Rows = GridRowsInitial
    .FixedRows = 1
    .FixedCols = 1
End With
    
setupTickerGridColumn 0, TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, "", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, "Name", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.currencyCode, TickerGridColumnWidths.CurrencyWidth, "Curr", True, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.bidSize, TickerGridColumnWidths.BidSizeWidth, "Bid size", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.bid, TickerGridColumnWidths.BidWidth, "Bid", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.ask, TickerGridColumnWidths.AskWidth, "Ask", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, "Ask size", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.trade, TickerGridColumnWidths.TradeWidth, "Last", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSizeWidth, "Last size", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.volume, TickerGridColumnWidths.VolumeWidth, "Volume", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, "Chg", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, "Chg %", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.highPrice, TickerGridColumnWidths.highWidth, "High", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.lowPrice, TickerGridColumnWidths.LowWidth, "Low", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.closePrice, TickerGridColumnWidths.CloseWidth, "Close", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.openInterest, TickerGridColumnWidths.openInterestWidth, "Open interest", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn 0, TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, "Description", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.symbol, TickerGridColumnWidths.SymbolWidth, "Symbol", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.sectype, TickerGridColumnWidths.SecTypeWidth, "Sec Type", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.expiry, TickerGridColumnWidths.ExpiryWidth, "Expiry", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.exchange, TickerGridColumnWidths.ExchangeWidth, "Exchange", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, "Right", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridColumns.strike, TickerGridColumnWidths.StrikeWidth, "Strike", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter

setupColumnMap TickerGridColumns.MaxColumn

End Sub

Private Sub setupSummaryTickerGrid()
With TickerGrid
    .RowBackColorEven = CellBackColorEven
    .RowBackColorOdd = CellBackColorOdd
    .AllowUserReordering = TwGridReorderBoth
    .AllowBigSelection = True
    .AllowUserResizing = TwGridResizeBoth
    .RowSizingMode = TwGridRowSizeAll
    .FillStyle = TwGridFillRepeat
    .FocusRect = TwGridFocusNone
    .HighLight = TwGridHighlightNever
    
    .Cols = 2
    .Rows = GridRowsInitial
    .FixedRows = 1
    .FixedCols = 1
End With
    
setupTickerGridColumn 0, TickerGridSummaryColumns.Selector, TickerGridSummaryColumnWidths.SelectorWidth, "", True, TWControls10.AlignmentSettings.TwGridAlignCenterBottom
setupTickerGridColumn 0, TickerGridSummaryColumns.TickerName, TickerGridSummaryColumnWidths.NameWidth, "Name", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.bidSize, TickerGridSummaryColumnWidths.BidSizeWidth, "Bid size", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.bid, TickerGridSummaryColumnWidths.BidWidth, "Bid", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.ask, TickerGridSummaryColumnWidths.AskWidth, "Ask", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.AskSize, TickerGridSummaryColumnWidths.AskSizeWidth, "Ask size", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.trade, TickerGridSummaryColumnWidths.TradeWidth, "Last", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.TradeSize, TickerGridSummaryColumnWidths.TradeSizeWidth, "Last size", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.volume, TickerGridSummaryColumnWidths.VolumeWidth, "Volume", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.Change, TickerGridSummaryColumnWidths.ChangeWidth, "Change", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.ChangePercent, TickerGridSummaryColumnWidths.ChangePercentWidth, "Change %", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn 0, TickerGridSummaryColumns.openInterest, TickerGridSummaryColumnWidths.openInterestWidth, "Open interest", False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter

setupColumnMap TickerGridSummaryColumns.MaxSummaryColumn

End Sub

Private Sub setupTickerGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As TWControls10.AlignmentSettings)
    
Dim lColumnWidth As Long

With TickerGrid
    .row = rowNumber
    
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .colWidth(columnNumber) = 0
    End If
    
    .ColData(columnNumber) = columnNumber
    
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    .colWidth(columnNumber) = lColumnWidth
    
    .ColAlignment(columnNumber) = align
    .ColAlignmentFixed(columnNumber) = align
    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With
End Sub
                
Private Sub stopTicker( _
                ByVal index As Long)
Dim lTicker As ticker

Set lTicker = mTickerTable(index).theTicker
lTicker.removeQuoteListener Me
lTicker.removePriceChangeListener Me

removeTicker index

lTicker.stopTicker
End Sub

Private Sub toggleRowSelection( _
                ByVal row As Long)
If isRowSelected(row) Then
    deselectRow row
Else
    selectRow row
End If
End Sub
                


