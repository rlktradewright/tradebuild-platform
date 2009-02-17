VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#30.0#0"; "TWControls10.ocx"
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

Event Click() 'MappingInfo=TickerGrid,TickerGrid,-1,Click
Event ColMoved(ByVal fromCol As Long, ByVal toCol As Long) 'MappingInfo=TickerGrid,TickerGrid,-1,ColMoved
Event ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean) 'MappingInfo=TickerGrid,TickerGrid,-1,ColMoving
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseUp
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseMove
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseDown
Event RowMoved(ByVal fromRow As Long, ByVal toRow As Long) 'MappingInfo=TickerGrid,TickerGrid,-1,RowMoved
Event RowMoving(ByVal fromRow As Long, ByVal toRow As Long, Cancel As Boolean) 'MappingInfo=TickerGrid,TickerGrid,-1,RowMoving
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyUp
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyPress
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyDown
Event DblClick() 'MappingInfo=TickerGrid,TickerGrid,-1,DblClick

Event TickerStarted(ByVal row As Long)

'@================================================================================
' Constants
'@================================================================================

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
    openPrice
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
    OpenWidth = 9
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

Private mPositiveChangeBackColor As OLE_COLOR
Private mPositiveChangeForeColor As OLE_COLOR
Private mNegativeChangeBackColor As OLE_COLOR
Private mNegativeChangeForeColor As OLE_COLOR

Private mIncreasedValueColor As OLE_COLOR
Private mDecreasedValueColor As OLE_COLOR

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

End Sub

Private Sub UserControl_InitProperties()
PositiveChangeBackColor = CPositiveChangeBackColor
PositiveChangeForeColor = CPositiveChangeForeColor
NegativeChangeBackColor = CNegativeChangeBackColor
NegativeChangeForeColor = CNegativeChangeForeColor
IncreasedValueColor = CIncreasedValueColor
DecreasedValueColor = CDecreasedValueColor
RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
AllowUserReordering = TwGridReorderBoth
TickerGrid.AllowBigSelection = True
AllowUserResizing = TwGridResizeBoth
RowSizingMode = TwGridRowSizeAll
FillStyle = TwGridFillRepeat
FocusRect = TwGridFocusNone
HighLight = TwGridHighlightNever
    
Cols = 2
Rows = GridRowsInitial
FixedRows = 1
FixedCols = 1

setupDefaultTickerGrid

End Sub

Private Sub UserControl_Resize()
TickerGrid.Left = 0
TickerGrid.Top = 0
TickerGrid.Width = UserControl.Width
TickerGrid.Height = UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim index As Integer

PositiveChangeBackColor = PropBag.ReadProperty("PositiveChangeBackColor", CPositiveChangeBackColor)
PositiveChangeForeColor = PropBag.ReadProperty("PositiveChangeForeColor", CPositiveChangeForeColor)
NegativeChangeBackColor = PropBag.ReadProperty("NegativeChangeBackColor", CNegativeChangeBackColor)
NegativeChangeForeColor = PropBag.ReadProperty("NegativeChangeForeColor", CNegativeChangeForeColor)
IncreasedValueColor = PropBag.ReadProperty("IncreasedValueColor", CIncreasedValueColor)
DecreasedValueColor = PropBag.ReadProperty("DecreasedValueColor", CDecreasedValueColor)
    
TickerGrid.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", TwGridResizeBoth)
TickerGrid.AllowUserReordering = PropBag.ReadProperty("AllowUserReordering", TwGridReorderBoth)
TickerGrid.BackColorSel = PropBag.ReadProperty("BackColorSel", -2147483635)
TickerGrid.BackColorFixed = PropBag.ReadProperty("BackColorFixed", -2147483633)
TickerGrid.BackColorBkg = PropBag.ReadProperty("BackColorBkg", -2147483636)
TickerGrid.backColor = PropBag.ReadProperty("BackColor", &H80000005)
TickerGrid.SelectionMode = PropBag.ReadProperty("SelectionMode", 2)
TickerGrid.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)
TickerGrid.RowSizingMode = PropBag.ReadProperty("RowSizingMode", TwGridRowSizeAll)
TickerGrid.Rows = PropBag.ReadProperty("Rows", 2)
TickerGrid.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
TickerGrid.RowBackColorOdd = PropBag.ReadProperty("RowBackColorOdd", CRowBackColorOdd)
TickerGrid.RowBackColorEven = PropBag.ReadProperty("RowBackColorEven", CRowBackColorEven)
TickerGrid.HighLight = PropBag.ReadProperty("HighLight", TwGridHighlightNever)
TickerGrid.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
TickerGrid.GridColorFixed = PropBag.ReadProperty("GridColorFixed", 12632256)
TickerGrid.GridColor = PropBag.ReadProperty("GridColor", -2147483631)
TickerGrid.ForeColorSel = PropBag.ReadProperty("ForeColorSel", -2147483634)
TickerGrid.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", -2147483630)
TickerGrid.foreColor = PropBag.ReadProperty("ForeColor", &H80000008)
TickerGrid.FocusRect = PropBag.ReadProperty("FocusRect", TwGridFocusNone)
TickerGrid.FixedRows = PropBag.ReadProperty("FixedRows", 1)
TickerGrid.FixedCols = PropBag.ReadProperty("FixedCols", 1)
TickerGrid.FillStyle = PropBag.ReadProperty("FillStyle", TwGridFillRepeat)
TickerGrid.Cols = PropBag.ReadProperty("Cols", 2)

setupDefaultTickerGrid

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim index As Integer

    Call PropBag.WriteProperty("AllowUserResizing", TickerGrid.AllowUserResizing, 3)
    Call PropBag.WriteProperty("AllowUserReordering", TickerGrid.AllowUserReordering, 0)
    Call PropBag.WriteProperty("BackColorSel", TickerGrid.BackColorSel, -2147483635)
    Call PropBag.WriteProperty("BackColorFixed", TickerGrid.BackColorFixed, -2147483633)
    Call PropBag.WriteProperty("BackColorBkg", TickerGrid.BackColorBkg, -2147483636)
    Call PropBag.WriteProperty("BackColor", TickerGrid.backColor, &H80000005)
    Call PropBag.WriteProperty("SelectionMode", TickerGrid.SelectionMode, 2)
    Call PropBag.WriteProperty("ScrollBars", TickerGrid.ScrollBars, 3)
    Call PropBag.WriteProperty("RowSizingMode", TickerGrid.RowSizingMode, 0)
    Call PropBag.WriteProperty("Rows", TickerGrid.Rows, 2)
    Call PropBag.WriteProperty("RowHeightMin", TickerGrid.RowHeightMin, 0)
    Call PropBag.WriteProperty("RowBackColorOdd", TickerGrid.RowBackColorOdd, 0)
    Call PropBag.WriteProperty("RowBackColorEven", TickerGrid.RowBackColorEven, 0)
    Call PropBag.WriteProperty("HighLight", TickerGrid.HighLight, 1)
    Call PropBag.WriteProperty("GridLineWidth", TickerGrid.GridLineWidth, 1)
    Call PropBag.WriteProperty("GridColorFixed", TickerGrid.GridColorFixed, 12632256)
    Call PropBag.WriteProperty("GridColor", TickerGrid.GridColor, -2147483631)
    Call PropBag.WriteProperty("ForeColorSel", TickerGrid.ForeColorSel, -2147483634)
    Call PropBag.WriteProperty("ForeColorFixed", TickerGrid.ForeColorFixed, -2147483630)
    Call PropBag.WriteProperty("ForeColor", TickerGrid.foreColor, &H80000008)
    Call PropBag.WriteProperty("FocusRect", TickerGrid.FocusRect, 1)
    Call PropBag.WriteProperty("FixedRows", TickerGrid.FixedRows, 1)
    Call PropBag.WriteProperty("FixedCols", TickerGrid.FixedCols, 1)
    Call PropBag.WriteProperty("FillStyle", TickerGrid.FillStyle, 0)
    Call PropBag.WriteProperty("Cols", TickerGrid.Cols, 2)
End Sub

'@================================================================================
' PriceChangeListener Interface Members
'@================================================================================

Private Sub PriceChangeListener_Change(ev As PriceChangeEvent)
Dim lTicker As ticker
Set lTicker = ev.Source

TickerGrid.row = getTickerGridRow(lTicker)
TickerGrid.col = mColumnMap(TickerGridColumns.Change)
TickerGrid.Text = ev.ChangeString
If ev.Change >= 0 Then
    TickerGrid.CellBackColor = mPositiveChangeBackColor
    TickerGrid.CellForeColor = mPositiveChangeForeColor
Else
    TickerGrid.CellBackColor = mNegativeChangeBackColor
    TickerGrid.CellForeColor = mNegativeChangeForeColor
End If
incrementEventCount

TickerGrid.col = mColumnMap(TickerGridColumns.ChangePercent)
TickerGrid.Text = Format(ev.ChangePercent, "0.0")
If ev.ChangePercent >= 0 Then
    TickerGrid.CellBackColor = mPositiveChangeBackColor
    TickerGrid.CellForeColor = mPositiveChangeForeColor
Else
    TickerGrid.CellBackColor = mNegativeChangeBackColor
    TickerGrid.CellForeColor = mNegativeChangeForeColor
End If

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

Private Sub QuoteListener_sessionOpen(ev As TradeBuild26.QuoteEvent)

displayPrice ev, mColumnMap(TickerGridColumns.openPrice)

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

RaiseEvent ColMoved(fromCol, toCol)
End Sub

Private Sub TickerGrid_ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean)
    RaiseEvent ColMoving(fromCol, toCol, Cancel)
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
            selectTicker row
        Else
            toggleRowSelection row
        End If
    End If
Else
    deselectSelectedTickers
End If

RaiseEvent Click
End Sub

Private Sub TickerGrid_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TickerGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TickerGrid_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TickerGrid_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TickerGrid_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TickerGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub TickerGrid_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseUp(Button, Shift, X, Y)
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

RaiseEvent RowMoved(fromRow, toRow)
End Sub

Private Sub TickerGrid_RowMoving( _
                ByVal fromRow As Long, _
                ByVal toRow As Long, _
                Cancel As Boolean)
If toRow > mNextGridRowIndex Then
    Cancel = True
Else
    RaiseEvent RowMoving(fromRow, toRow, Cancel)
End If
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

index = getTickerIndex(lTicker)
    
Select Case ev.State
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
    TickerGrid.Text = IIf(lContract.expiryDate = 0, "", lContract.expiryDate)
    
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
    
    RaiseEvent TickerStarted(mTickerTable(index).tickerGridRow)
    
Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    ' if the ticker was stopped by the application via a call to Ticker.stopTicker (rather
    ' tha via this control), the entry will still be in the grid so remove it
    If Not mTickerTable(index).theTicker Is Nothing Then
        removeTicker index
    End If
End Select
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let PositiveChangeBackColor(ByVal value As OLE_COLOR)
mPositiveChangeBackColor = value
PropertyChanged "PositiveChangeBackColor"
End Property

Public Property Get PositiveChangeBackColor() As OLE_COLOR
PositiveChangeBackColor = mPositiveChangeBackColor
End Property

Public Property Let PositiveChangeForeColor(ByVal value As OLE_COLOR)
mPositiveChangeForeColor = value
PropertyChanged "PositiveChangeForeColor"
End Property

Public Property Get PositiveChangeForeColor() As OLE_COLOR
PositiveChangeForeColor = mPositiveChangeForeColor
End Property

Public Property Let NegativeChangeBackColor(ByVal value As OLE_COLOR)
mNegativeChangeBackColor = value
PropertyChanged "NegativeChangeBackColor"
End Property

Public Property Get NegativeChangeBackColor() As OLE_COLOR
NegativeChangeBackColor = mNegativeChangeBackColor
End Property

Public Property Let NegativeChangeForeColor(ByVal value As OLE_COLOR)
mNegativeChangeForeColor = value
PropertyChanged "NegativeChangeForeColor"
End Property

Public Property Get NegativeChangeForeColor() As OLE_COLOR
NegativeChangeForeColor = mNegativeChangeForeColor
End Property

Public Property Let IncreasedValueColor(ByVal value As OLE_COLOR)
mIncreasedValueColor = value
PropertyChanged "IncreasedValueColor"
End Property

Public Property Get IncreasedValueColor() As OLE_COLOR
IncreasedValueColor = mIncreasedValueColor
End Property

Public Property Let DecreasedValueColor(ByVal value As OLE_COLOR)
mDecreasedValueColor = value
PropertyChanged "DecreasedValueColor"
End Property

Public Property Get DecreasedValueColor() As OLE_COLOR
DecreasedValueColor = mDecreasedValueColor
End Property

Public Property Get AllowUserResizing() As AllowUserResizeSettings
    AllowUserResizing = TickerGrid.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
    TickerGrid.AllowUserResizing = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"
End Property

Public Property Get AllowUserReordering() As AllowUserReorderSettings
    AllowUserReordering = TickerGrid.AllowUserReordering
End Property

Public Property Let AllowUserReordering(ByVal New_AllowUserReordering As AllowUserReorderSettings)
    TickerGrid.AllowUserReordering = New_AllowUserReordering
    PropertyChanged "AllowUserReordering"
End Property

Public Property Get BackColorSel() As OLE_COLOR
    BackColorSel = TickerGrid.BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
    TickerGrid.BackColorSel = New_BackColorSel
    PropertyChanged "BackColorSel"
End Property

Public Property Get BackColorFixed() As OLE_COLOR
    BackColorFixed = TickerGrid.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    TickerGrid.BackColorFixed = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
    BackColorBkg = TickerGrid.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    TickerGrid.BackColorBkg = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

Public Property Get backColor() As OLE_COLOR
    backColor = TickerGrid.backColor
End Property

Public Property Let backColor(ByVal New_BackColor As OLE_COLOR)
    TickerGrid.backColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get TopRow() As Long
    TopRow = TickerGrid.TopRow
End Property

Public Property Let TopRow(ByVal New_TopRow As Long)
    TickerGrid.TopRow = New_TopRow
    PropertyChanged "TopRow"
End Property

Public Property Get TextStyleFixed() As TextStyleSettings
    TextStyleFixed = TickerGrid.TextStyleFixed
End Property

Public Property Let TextStyleFixed(ByVal New_TextStyleFixed As TextStyleSettings)
    TickerGrid.TextStyleFixed = New_TextStyleFixed
    PropertyChanged "TextStyleFixed"
End Property

Public Property Get TextStyle() As TextStyleSettings
    TextStyle = TickerGrid.TextStyle
End Property

Public Property Let TextStyle(ByVal New_TextStyle As TextStyleSettings)
    TickerGrid.TextStyle = New_TextStyle
    PropertyChanged "TextStyle"
End Property

Public Property Get TextMatrix(ByVal row As Long, ByVal col As Long) As String
    TextMatrix = TickerGrid.TextMatrix(row, col)
End Property

Public Property Let TextMatrix(ByVal row As Long, ByVal col As Long, ByVal New_TextMatrix As String)
    TickerGrid.TextMatrix(row, col) = New_TextMatrix
    PropertyChanged "TextMatrix"
End Property

Public Property Get TextArray(ByVal index As Long) As String
    TextArray = TickerGrid.TextArray(index)
End Property

Public Property Let TextArray(ByVal index As Long, ByVal New_TextArray As String)
    TickerGrid.TextArray(index) = New_TextArray
    PropertyChanged "TextArray"
End Property

Public Property Get Text() As String
    Text = TickerGrid.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    TickerGrid.Text = New_Text
    PropertyChanged "Text"
End Property

Public Property Get SelectionMode() As SelectionModeSettings
    SelectionMode = TickerGrid.SelectionMode
End Property

Public Property Let SelectionMode(ByVal New_SelectionMode As SelectionModeSettings)
    TickerGrid.SelectionMode = New_SelectionMode
    PropertyChanged "SelectionMode"
End Property

Public Property Get ScrollBars() As ScrollBarsSettings
    ScrollBars = TickerGrid.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    TickerGrid.ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

Public Property Get RowSizingMode() As RowSizingSettings
    RowSizingMode = TickerGrid.RowSizingMode
End Property

Public Property Let RowSizingMode(ByVal New_RowSizingMode As RowSizingSettings)
    TickerGrid.RowSizingMode = New_RowSizingMode
    PropertyChanged "RowSizingMode"
End Property

Public Property Get rowSel() As Long
    rowSel = TickerGrid.rowSel
End Property

Public Property Let rowSel(ByVal New_RowSel As Long)
    TickerGrid.rowSel = New_RowSel
    PropertyChanged "RowSel"
End Property

Public Property Get Rows() As Long
    Rows = TickerGrid.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    TickerGrid.Rows = New_Rows
    PropertyChanged "Rows"
End Property

Public Property Let RowPosition(ByVal index As Long, ByVal New_RowPosition As Long)
    TickerGrid.RowPosition(index) = New_RowPosition
    PropertyChanged "RowPosition"
End Property

Public Property Get RowPos(ByVal index As Long) As Long
    RowPos = TickerGrid.RowPos(index)
End Property

Public Property Get RowIsVisible(ByVal index As Long) As Boolean
    RowIsVisible = TickerGrid.RowIsVisible(index)
End Property

Public Property Get RowHeightMin() As Long
    RowHeightMin = TickerGrid.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    TickerGrid.RowHeightMin = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

Public Property Get rowHeight(ByVal index As Long) As Long
    rowHeight = TickerGrid.rowHeight(index)
End Property

Public Property Let rowHeight(ByVal index As Long, ByVal New_rowHeight As Long)
    TickerGrid.rowHeight(index) = New_rowHeight
    PropertyChanged "rowHeight"
End Property

Public Property Get rowdata(ByVal index As Long) As Long
    rowdata = TickerGrid.rowdata(index)
End Property

Public Property Let rowdata(ByVal index As Long, ByVal New_RowData As Long)
    TickerGrid.rowdata(index) = New_RowData
    PropertyChanged "RowData"
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
    RowBackColorOdd = TickerGrid.RowBackColorOdd
End Property

Public Property Let RowBackColorOdd(ByVal New_RowBackColorOdd As OLE_COLOR)
    TickerGrid.RowBackColorOdd = New_RowBackColorOdd
    PropertyChanged "RowBackColorOdd"
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
    RowBackColorEven = TickerGrid.RowBackColorEven
End Property

Public Property Let RowBackColorEven(ByVal New_RowBackColorEven As OLE_COLOR)
    TickerGrid.RowBackColorEven = New_RowBackColorEven
    PropertyChanged "RowBackColorEven"
End Property

Public Property Get row() As Long
    row = TickerGrid.row
End Property

Public Property Let row(ByVal New_row As Long)
    TickerGrid.row = New_row
    PropertyChanged "row"
End Property

Public Property Get Redraw() As Boolean
    Redraw = TickerGrid.Redraw
End Property

Public Property Let Redraw(ByVal New_Redraw As Boolean)
    TickerGrid.Redraw = New_Redraw
    PropertyChanged "Redraw"
End Property

Public Property Get MouseRow() As Long
    MouseRow = TickerGrid.MouseRow
End Property

Public Property Get MouseCol() As Long
    MouseCol = TickerGrid.MouseCol
End Property

Public Property Get LeftCol() As Long
    LeftCol = TickerGrid.LeftCol
End Property

Public Property Let LeftCol(ByVal New_LeftCol As Long)
    TickerGrid.LeftCol = New_LeftCol
    PropertyChanged "LeftCol"
End Property

Public Property Get hWnd() As Long
    hWnd = TickerGrid.hWnd
End Property

Public Property Get HighLight() As HighLightSettings
    HighLight = TickerGrid.HighLight
End Property

Public Property Let HighLight(ByVal New_HighLight As HighLightSettings)
    TickerGrid.HighLight = New_HighLight
    PropertyChanged "HighLight"
End Property

Public Property Get GridLineWidth() As Long
    GridLineWidth = TickerGrid.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Long)
    TickerGrid.GridLineWidth = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridLinesFixed() As GridLineSettings
    GridLinesFixed = TickerGrid.GridLinesFixed
End Property

Public Property Let GridLinesFixed(ByVal New_GridLinesFixed As GridLineSettings)
    TickerGrid.GridLinesFixed = New_GridLinesFixed
    PropertyChanged "GridLinesFixed"
End Property

Public Property Get GridLines() As GridLineSettings
    GridLines = TickerGrid.GridLines
End Property

Public Property Let GridLines(ByVal New_GridLines As GridLineSettings)
    TickerGrid.GridLines = New_GridLines
    PropertyChanged "GridLines"
End Property

Public Property Get GridColorFixed() As OLE_COLOR
    GridColorFixed = TickerGrid.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
    TickerGrid.GridColorFixed = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

Public Property Get GridColor() As OLE_COLOR
    GridColor = TickerGrid.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    TickerGrid.GridColor = New_GridColor
    PropertyChanged "GridColor"
End Property

Public Function getRowFromY(ByVal Y As Long) As Long
    getRowFromY = TickerGrid.getRowFromY(Y)
End Function

Public Function getColFromX(ByVal X As Long) As Long
    getColFromX = TickerGrid.getColFromX(X)
End Function

Public Property Get FormatString() As String
    FormatString = TickerGrid.FormatString
End Property

Public Property Let FormatString(ByVal New_FormatString As String)
    TickerGrid.FormatString = New_FormatString
    PropertyChanged "FormatString"
End Property

Public Property Get ForeColorSel() As OLE_COLOR
    ForeColorSel = TickerGrid.ForeColorSel
End Property

Public Property Let ForeColorSel(ByVal New_ForeColorSel As OLE_COLOR)
    TickerGrid.ForeColorSel = New_ForeColorSel
    PropertyChanged "ForeColorSel"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
    ForeColorFixed = TickerGrid.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    TickerGrid.ForeColorFixed = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property

Public Property Get foreColor() As OLE_COLOR
    foreColor = TickerGrid.foreColor
End Property

Public Property Let foreColor(ByVal New_ForeColor As OLE_COLOR)
    TickerGrid.foreColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get FontWidthFixed() As Single
    FontWidthFixed = TickerGrid.FontWidthFixed
End Property

Public Property Let FontWidthFixed(ByVal New_FontWidthFixed As Single)
    TickerGrid.FontWidthFixed = New_FontWidthFixed
    PropertyChanged "FontWidthFixed"
End Property

Public Property Get FontWidth() As Single
    FontWidth = TickerGrid.FontWidth
End Property

Public Property Let FontWidth(ByVal New_FontWidth As Single)
    TickerGrid.FontWidth = New_FontWidth
    PropertyChanged "FontWidth"
End Property

Public Property Get FontFixed() As Font
    Set FontFixed = TickerGrid.FontFixed
End Property

Public Property Get Font() As Font
    Set Font = TickerGrid.Font
End Property

Public Property Get FocusRect() As FocusRectSettings
    FocusRect = TickerGrid.FocusRect
End Property

Public Property Let FocusRect(ByVal New_FocusRect As FocusRectSettings)
    TickerGrid.FocusRect = New_FocusRect
    PropertyChanged "FocusRect"
End Property

Public Property Get FixedRows() As Long
    FixedRows = TickerGrid.FixedRows
End Property

Public Property Let FixedRows(ByVal New_FixedRows As Long)
    TickerGrid.FixedRows = New_FixedRows
    PropertyChanged "FixedRows"
End Property

Public Property Get FixedCols() As Long
    FixedCols = TickerGrid.FixedCols
End Property

Public Property Let FixedCols(ByVal New_FixedCols As Long)
    TickerGrid.FixedCols = New_FixedCols
    PropertyChanged "FixedCols"
End Property

Public Property Get FillStyle() As FillStyleSettings
    FillStyle = TickerGrid.FillStyle
End Property

Public Property Let FillStyle(ByVal New_FillStyle As FillStyleSettings)
    TickerGrid.FillStyle = New_FillStyle
    PropertyChanged "FillStyle"
End Property

Public Property Get colWidth(ByVal index As Long) As Long
    colWidth = TickerGrid.colWidth(index)
End Property

Public Property Let colWidth(ByVal index As Long, ByVal New_colWidth As Long)
    TickerGrid.colWidth(index) = New_colWidth
    PropertyChanged "colWidth"
End Property

Public Property Get colSel() As Long
    colSel = TickerGrid.colSel
End Property

Public Property Let colSel(ByVal New_ColSel As Long)
    TickerGrid.colSel = New_ColSel
    PropertyChanged "ColSel"
End Property

Public Property Get Cols() As Long
    Cols = TickerGrid.Cols
End Property

Public Property Let Cols(ByVal New_Cols As Long)
    TickerGrid.Cols = New_Cols
    PropertyChanged "Cols"
End Property

Public Property Let ColPosition(ByVal index As Long, ByVal New_ColPosition As Long)
    TickerGrid.ColPosition(index) = New_ColPosition
    PropertyChanged "ColPosition"
End Property

Public Property Get ColPos(ByVal index As Long) As Long
    ColPos = TickerGrid.ColPos(index)
End Property

Public Property Get ColIsVisible(ByVal index As Long) As Boolean
    ColIsVisible = TickerGrid.ColIsVisible(index)
End Property

Public Property Get ColData(ByVal index As Long) As Long
    ColData = TickerGrid.ColData(index)
End Property

Public Property Let ColData(ByVal index As Long, ByVal New_ColData As Long)
    TickerGrid.ColData(index) = New_ColData
    PropertyChanged "ColData"
End Property

Public Property Get ColAlignmentFixed(ByVal index As Long) As AlignmentSettings
    ColAlignmentFixed = TickerGrid.ColAlignmentFixed(index)
End Property

Public Property Let ColAlignmentFixed(ByVal index As Long, ByVal New_ColAlignmentFixed As AlignmentSettings)
    TickerGrid.ColAlignmentFixed(index) = New_ColAlignmentFixed
    PropertyChanged "ColAlignmentFixed"
End Property

Public Property Get ColAlignment(ByVal index As Long) As AlignmentSettings
    ColAlignment = TickerGrid.ColAlignment(index)
End Property

Public Property Let ColAlignment(ByVal index As Long, ByVal New_ColAlignment As AlignmentSettings)
    TickerGrid.ColAlignment(index) = New_ColAlignment
    PropertyChanged "ColAlignment"
End Property

Public Property Get col() As Long
    col = TickerGrid.col
End Property

Public Property Let col(ByVal New_col As Long)
    TickerGrid.col = New_col
    PropertyChanged "col"
End Property

Public Property Get cellWidth() As Long
    cellWidth = TickerGrid.cellWidth
End Property

Public Property Get celltop() As Long
    celltop = TickerGrid.celltop
End Property

Public Property Get CellTextStyle() As TextStyleSettings
    CellTextStyle = TickerGrid.CellTextStyle
End Property

Public Property Let CellTextStyle(ByVal New_CellTextStyle As TextStyleSettings)
    TickerGrid.CellTextStyle = New_CellTextStyle
    PropertyChanged "CellTextStyle"
End Property

Public Property Get CellPictureAlignment() As AlignmentSettings
    CellPictureAlignment = TickerGrid.CellPictureAlignment
End Property

Public Property Let CellPictureAlignment(ByVal New_CellPictureAlignment As AlignmentSettings)
    TickerGrid.CellPictureAlignment = New_CellPictureAlignment
    PropertyChanged "CellPictureAlignment"
End Property

Public Property Get CellPicture() As Picture
    Set CellPicture = TickerGrid.CellPicture
End Property

Public Property Get cellLeft() As Long
    cellLeft = TickerGrid.cellLeft
End Property

Public Property Get cellHeight() As Long
    cellHeight = TickerGrid.cellHeight
End Property

Public Property Get CellForeColor() As OLE_COLOR
    CellForeColor = TickerGrid.CellForeColor
End Property

Public Property Let CellForeColor(ByVal New_CellForeColor As OLE_COLOR)
    TickerGrid.CellForeColor = New_CellForeColor
    PropertyChanged "CellForeColor"
End Property

Public Property Get CellFontUnderline() As Boolean
    CellFontUnderline = TickerGrid.CellFontUnderline
End Property

Public Property Let CellFontUnderline(ByVal New_CellFontUnderline As Boolean)
    TickerGrid.CellFontUnderline = New_CellFontUnderline
    PropertyChanged "CellFontUnderline"
End Property

Public Property Get CellFontStrikeThrough() As Boolean
    CellFontStrikeThrough = TickerGrid.CellFontStrikeThrough
End Property

Public Property Let CellFontStrikeThrough(ByVal New_CellFontStrikeThrough As Boolean)
    TickerGrid.CellFontStrikeThrough = New_CellFontStrikeThrough
    PropertyChanged "CellFontStrikeThrough"
End Property

Public Property Get CellFontSize() As Single
    CellFontSize = TickerGrid.CellFontSize
End Property

Public Property Let CellFontSize(ByVal New_CellFontSize As Single)
    TickerGrid.CellFontSize = New_CellFontSize
    PropertyChanged "CellFontSize"
End Property

Public Property Get CellFontName() As String
    CellFontName = TickerGrid.CellFontName
End Property

Public Property Let CellFontName(ByVal New_CellFontName As String)
    TickerGrid.CellFontName = New_CellFontName
    PropertyChanged "CellFontName"
End Property

Public Property Get CellFontItalic() As Boolean
    CellFontItalic = TickerGrid.CellFontItalic
End Property

Public Property Let CellFontItalic(ByVal New_CellFontItalic As Boolean)
    TickerGrid.CellFontItalic = New_CellFontItalic
    PropertyChanged "CellFontItalic"
End Property

Public Property Get CellFontBold() As Boolean
    CellFontBold = TickerGrid.CellFontBold
End Property

Public Property Let CellFontBold(ByVal New_CellFontBold As Boolean)
    TickerGrid.CellFontBold = New_CellFontBold
    PropertyChanged "CellFontBold"
End Property

Public Property Get CellBackColor() As OLE_COLOR
    CellBackColor = TickerGrid.CellBackColor
End Property

Public Property Let CellBackColor(ByVal New_CellBackColor As OLE_COLOR)
    TickerGrid.CellBackColor = New_CellBackColor
    PropertyChanged "CellBackColor"
End Property

Public Property Get CellAlignment() As AlignmentSettings
    CellAlignment = TickerGrid.CellAlignment
End Property

Public Property Let CellAlignment(ByVal New_CellAlignment As AlignmentSettings)
    TickerGrid.CellAlignment = New_CellAlignment
    PropertyChanged "CellAlignment"
End Property

Public Property Get SelectedTickers() As SelectedTickers
Dim index As Long

Set SelectedTickers = New SelectedTickers

index = mFirstSelected

Do While index <> 0
    SelectedTickers.add mTickerTable(index).theTicker
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

Public Sub deselectTicker( _
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

'Public Sub ExtendSelection(ByVal row As Long, ByVal col As Long)
'    TickerGrid.ExtendSelection row, col
'End Sub

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

Public Sub InvertCellColors()
TickerGrid.InvertCellColors
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
    selectTicker i
Next
End Sub

Public Sub selectTicker( _
                ByVal row As Long)
Dim index As Long
index = TickerGrid.rowdata(row)
If Not mTickerTable(index).theTicker Is Nothing Then
    mTickerTable(index).nextSelected = mFirstSelected
    mTickerTable(index).prevSelected = -1
    mTickerTable(mFirstSelected).prevSelected = index
    mFirstSelected = index
    highlightRow row
End If
End Sub

Public Sub setCellAlignment(ByVal row As Long, ByVal col As Long, pAlign As AlignmentSettings)
TickerGrid.setCellAlignment row, col, pAlign
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

'Private Sub deselectRow( _
'                ByVal row As Long)
'deselectTicker TickerGrid.rowdata(row)
'End Sub

Private Sub displayPrice( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = getTickerGridRow(lTicker)
TickerGrid.col = col
TickerGrid.Text = ev.priceString
If ev.priceChange = ValueChangeUp Then
    TickerGrid.CellForeColor = mIncreasedValueColor
ElseIf ev.priceChange = ValueChangeDown Then
    TickerGrid.CellForeColor = mDecreasedValueColor
End If

incrementEventCount
End Sub

Private Sub displaySize( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As ticker
Set lTicker = ev.Source
TickerGrid.row = getTickerGridRow(lTicker)
TickerGrid.col = col
TickerGrid.Text = ev.size
If ev.sizeChange = ValueChangeUp Then
    TickerGrid.CellForeColor = mIncreasedValueColor
ElseIf ev.sizeChange = ValueChangeDown Then
    TickerGrid.CellForeColor = mDecreasedValueColor
End If

incrementEventCount
End Sub

Private Function getTickerGridRow( _
                ByVal pTicker As ticker) As Long
getTickerGridRow = mTickerTable(getTickerIndex(pTicker)).tickerGridRow
End Function

Private Function getTickerIndex( _
                ByVal pTicker As ticker) As Long
' allow for the fact that the first tickertable entry is not used - it is the
' terminator of the selected entries chain
getTickerIndex = pTicker.Handle + 1
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

'Private Sub SelectRow( _
'                ByVal row As Long)
'selectTicker TickerGrid.rowdata(row)
'End Sub

Private Sub setupColumnMap( _
                    ByVal maxIndex As Long)
Dim i As Long
ReDim mColumnMap(maxIndex) As Long
For i = 0 To UBound(mColumnMap)
    mColumnMap(i) = i
Next
End Sub

Private Sub setupDefaultTickerGrid()

'With TickerGrid
'    .RowBackColorEven = CellBackColorEven
'    .RowBackColorOdd = CellBackColorOdd
'    .AllowUserReordering = TwGridReorderBoth
'    .AllowBigSelection = True
'    .AllowUserResizing = TwGridResizeBoth
'    .RowSizingMode = TwGridRowSizeAll
'    .FillStyle = TwGridFillRepeat
'    .FocusRect = TwGridFocusNone
'    .HighLight = TwGridHighlightNever
'
'    .Cols = 2
'    .Rows = GridRowsInitial
'    .FixedRows = 1
'    .FixedCols = 1
'End With
    
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
setupTickerGridColumn 0, TickerGridColumns.openPrice, TickerGridColumnWidths.OpenWidth, "Open", False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
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
    .RowBackColorEven = CRowBackColorEven
    .RowBackColorOdd = CRowBackColorOdd
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
    deselectTicker row
Else
    selectTicker row
End If
End Sub

