VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.UserControl TickerGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox FontPicture 
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   1515
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin TWControls40.TWGrid TickerGrid 
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

Implements IDeferredAction
Implements IErrorListener
Implements IQuoteListener
Implements IPriceChangeListener
Implements IThemeable
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Event ColMoved(ByVal fromCol As Long, ByVal toCol As Long)
Event ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean)
Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Event Error(ev As ErrorEventData)
Event ErroredTickerRemoved(ByVal pTicker As IMarketDataSource)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event RowMoved(ByVal pFromRow As Long, ByVal pToRow As Long)
Event RowMoving(ByVal pFromRow As Long, ByVal pToRow As Long, Cancel As Boolean)
Event SelectionChanged(ByVal pRow1 As Long, ByVal pCol1 As Long, ByVal pRow2 As Long, ByVal pCol2 As Long)
Event TickerSelectionChanged()
Event TickerSymbolEntered(ByVal pSymbol As String, ByVal pPreferredRow As Long)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "TickerGrid"

Private Const ConfigSectionGrid                         As String = "Grid"
Private Const ConfigSectionTicker                       As String = "Ticker"
Private Const ConfigSectionTickers                      As String = "Tickers"

Private Const ConfigSettingRowIndex                     As String = "&RowIndex"

Private Const ConfigSettingPositiveChangeBackColor      As String = "&PositiveChangeBackColor"
Private Const ConfigSettingPositiveChangeForeColor      As String = "&PositiveChangeForeColor"
Private Const ConfigSettingNegativeChangeBackColor      As String = "&NegativeChangeBackColor"
Private Const ConfigSettingNegativeChangeForeColor      As String = "&NegativeChangeForeColor"
Private Const ConfigSettingIncreasedValueColor          As String = "&IncreasedValueColor"
Private Const ConfigSettingHighlightPriceChanges        As String = "&HighlightPriceChanges"
Private Const ConfigSettingDecreasedValueColor          As String = "&DecreasedValueColor"
Private Const ConfigSettingColumnMap                    As String = ".ColumnMap"

Private Const GridRowsInitial                           As Long = 100
Private Const GridRowsIncrement                         As Long = 50

Private Const TickerTableEntriesInitial                 As Long = 4
Private Const TickerTableEntriesGrowthFactor            As Long = 2

'@================================================================================
' Enums
'@================================================================================

Private Enum TickerGridColumns
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'NB: don't ever change the values of these items, since they might
    ' be persisted in the column map in the config file.
    ' Changes in the display order can be effected by changing the column
    ' map.
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Selector = 0
    TickerName = 1
    CurrencyCode = 2
    BidSize = 3
    Bid = 4
    Ask = 5
    AskSize = 6
    Trade = 7
    TradeSize = 8
    Volume = 9
    Change = 10
    ChangePercent = 11
    HighPrice = 12
    LowPrice = 13
    OpenPrice = 14
    ClosePrice = 15
    OpenInterest = 16
    Description = 17
    Symbol = 18
    secType = 19
    Expiry = 20
    Exchange = 21
    OptionRight = 22
    Strike = 23
    ErrorText = 24
    MaxColumn = 24
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
    HighWidth = 9
    LowWidth = 9
    OpenWidth = 9
    CloseWidth = 9
    OpenInterestWidth = 9
    DescriptionWidth = 20
    SymbolWidth = 5
    SecTypeWidth = 10
    ExpiryWidth = 10
    ExchangeWidth = 10
    OptionRightWidth = 5
    StrikeWidth = 9
    ErrorTextWidth = 30
End Enum

Private Enum TickerGridSummaryColumns
    Selector = 0
    TickerName = 1
    BidSize = 2
    Bid = 3
    Ask = 4
    AskSize = 5
    Trade = 6
    TradeSize = 7
    Volume = 8
    Change = 9
    ChangePercent = 10
    OpenInterest = 11
    ErrorText = 12
    MaxSummaryColumn = 12
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
    OpenInterestWidth
    ErrorTextWidth = 30
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type TickerTableEntry
    DataSource              As IMarketDataSource
    TickerGridRow           As Long
    FieldsHaveBeenSet       As Boolean
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mMarketDataManager                              As IMarketDataManager

Private mTickerTable()                                  As TickerTableEntry
Private mTickers                                        As New EnumerableCollection

Private mLetterWidth                                    As Single
Private mDigitWidth                                     As Single

Private mControlDown                                    As Boolean
Private mShiftDown                                      As Boolean
Private mAltDown                                        As Boolean

Private mColumnMap()                                    As Long

Private WithEvents mSelectedTickers                     As SelectedTickers
Attribute mSelectedTickers.VB_VarHelpID = -1

Private mPositiveChangeBackColor                        As OLE_COLOR
Private mPositiveChangeForeColor                        As OLE_COLOR
Private mNegativeChangeBackColor                        As OLE_COLOR
Private mNegativeChangeForeColor                        As OLE_COLOR

Private mIncreasedValueColor                            As OLE_COLOR
Private mDecreasedValueColor                            As OLE_COLOR

Private mConfig                                         As ConfigurationSection
Private mTickersConfigSection                           As ConfigurationSection

Private mHighlightPriceChanges                          As Boolean

Private mEnteringTickerSymbol                           As Boolean
Private mTickerSymbolRow                                As Long

Private mIteratingTickersConfig                         As Boolean

Private mTheme                                          As ITheme

Private WithEvents mRefreshQuotesTC                     As TaskController
Attribute mRefreshQuotesTC.VB_VarHelpID = -1
Private WithEvents mRefreshPriceChangeTC                As TaskController
Attribute mRefreshPriceChangeTC.VB_VarHelpID = -1

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
setupColumnMap TickerGridColumns.MaxColumn

calcAverageCharacterWidths UserControl.Font

Set mSelectedTickers = New SelectedTickers

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
Const ProcName As String = "UserControl_InitProperties"
On Error GoTo Err

BorderStyle = BorderStyleNone
PositiveChangeBackColor = CPositiveChangeBackColor
PositiveChangeForeColor = CPositiveChangeForeColor
NegativeChangeBackColor = CNegativeChangeBackColor
NegativeChangeForeColor = CNegativeChangeForeColor
IncreasedValueColor = CIncreasedValueColor
DecreasedValueColor = CDecreasedValueColor
RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
AllowUserReordering = TwGridReorderBoth
HighlightPriceChanges = True

TickerGrid.AllowBigSelection = True
AllowUserResizing = TwGridResizeBoth
RowSizingMode = TwGridRowSizeAll
FillStyle = TwGridFillRepeat
TickerGrid.SelectionMode = TwGridSelectionFree
TickerGrid.FocusRect = TwGridFocusBroken
TickerGrid.HighLight = TwGridHighlightWithFocus
TickerGrid.FontFixed = UserControl.Ambient.Font
TickerGrid.Font = UserControl.Ambient.Font
    
TickerGrid.Cols = 2
Rows = GridRowsInitial
TickerGrid.FixedRows = 1
TickerGrid.FixedCols = 1

TickerGrid.PopupScrollbars = True

'setupDefaultTickerGridColumns
'setupDefaultTickerGridHeaders

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

TickerGrid.Left = 0
TickerGrid.Top = 0
TickerGrid.Width = UserControl.Width
TickerGrid.Height = UserControl.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

PositiveChangeBackColor = PropBag.ReadProperty("PositiveChangeBackColor", CPositiveChangeBackColor)
PositiveChangeForeColor = PropBag.ReadProperty("PositiveChangeForeColor", CPositiveChangeForeColor)
NegativeChangeBackColor = PropBag.ReadProperty("NegativeChangeBackColor", CNegativeChangeBackColor)
NegativeChangeForeColor = PropBag.ReadProperty("NegativeChangeForeColor", CNegativeChangeForeColor)
IncreasedValueColor = PropBag.ReadProperty("IncreasedValueColor", CIncreasedValueColor)
DecreasedValueColor = PropBag.ReadProperty("DecreasedValueColor", CDecreasedValueColor)
HighlightPriceChanges = PropBag.ReadProperty("HighlightPriceChanges", True)

TickerGrid.AllowUserResizing = PropBag.ReadProperty("AllowUserResizing", TwGridResizeBoth)
TickerGrid.AllowUserReordering = PropBag.ReadProperty("AllowUserReordering", TwGridReorderBoth)
TickerGrid.BackColorSel = PropBag.ReadProperty("BackColorSel", -2147483635)
TickerGrid.BackColorFixed = PropBag.ReadProperty("BackColorFixed", -2147483633)
TickerGrid.BackColorBkg = PropBag.ReadProperty("BackColorBkg", -2147483636)
TickerGrid.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
TickerGrid.BorderStyle = PropBag.ReadProperty("BorderStyle", BorderStyleNone)
TickerGrid.ScrollBars = PropBag.ReadProperty("ScrollBars", TwGridScrollBarBoth)
TickerGrid.RowSizingMode = PropBag.ReadProperty("RowSizingMode", TwGridRowSizeAll)
TickerGrid.Rows = PropBag.ReadProperty("Rows", GridRowsInitial)
TickerGrid.RowHeightMin = PropBag.ReadProperty("RowHeightMin", 0)
TickerGrid.RowBackColorOdd = PropBag.ReadProperty("RowBackColorOdd", CRowBackColorOdd)
TickerGrid.RowBackColorEven = PropBag.ReadProperty("RowBackColorEven", CRowBackColorEven)
TickerGrid.GridLinesFixed = PropBag.ReadProperty("GridLinesFixed", TwGridGridFlat)
TickerGrid.GridLines = PropBag.ReadProperty("GridLines", TwGridGridNone)
TickerGrid.GridLineWidth = PropBag.ReadProperty("GridLineWidth", 1)
TickerGrid.GridColorFixed = PropBag.ReadProperty("GridColorFixed", 12632256)
TickerGrid.GridColor = PropBag.ReadProperty("GridColor", -2147483631)
TickerGrid.FontFixed = PropBag.ReadProperty("FontFixed", UserControl.Ambient.Font)
TickerGrid.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
TickerGrid.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", -2147483630)
TickerGrid.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)

TickerGrid.FocusRect = TwGridFocusBroken
TickerGrid.FixedRows = 1
TickerGrid.FixedCols = 1
TickerGrid.PopupScrollbars = True
TickerGrid.FillStyle = TwGridFillRepeat
TickerGrid.Cols = 2
TickerGrid.ForeColorSel = -2147483634
TickerGrid.HighLight = TwGridHighlightWithFocus
TickerGrid.SelectionMode = TwGridSelectionFree

'setupDefaultTickerGridColumns
'setupDefaultTickerGridHeaders

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

On Error Resume Next
Call PropBag.WriteProperty("PositiveChangeBackColor", PositiveChangeBackColor, CPositiveChangeBackColor)
Call PropBag.WriteProperty("PositiveChangeForeColor", PositiveChangeForeColor, CPositiveChangeForeColor)
Call PropBag.WriteProperty("NegativeChangeBackColor", NegativeChangeBackColor, CNegativeChangeBackColor)
Call PropBag.WriteProperty("NegativeChangeForeColor", NegativeChangeForeColor, CNegativeChangeForeColor)
Call PropBag.WriteProperty("IncreasedValueColor", IncreasedValueColor, CIncreasedValueColor)
Call PropBag.WriteProperty("DecreasedValueColor", DecreasedValueColor, CDecreasedValueColor)
Call PropBag.WriteProperty("HighlightPriceChanges", HighlightPriceChanges, True)

Call PropBag.WriteProperty("AllowUserResizing", TickerGrid.AllowUserResizing, 3)
Call PropBag.WriteProperty("AllowUserReordering", TickerGrid.AllowUserReordering, 0)
Call PropBag.WriteProperty("BackColorSel", TickerGrid.BackColorSel, -2147483635)
Call PropBag.WriteProperty("BackColorFixed", TickerGrid.BackColorFixed, -2147483633)
Call PropBag.WriteProperty("BackColorBkg", TickerGrid.BackColorBkg, -2147483636)
Call PropBag.WriteProperty("BackColor", TickerGrid.BackColor, &H80000005)
Call PropBag.WriteProperty("BorderStyle", TickerGrid.BorderStyle, BorderStyleNone)
Call PropBag.WriteProperty("SelectionMode", TickerGrid.SelectionMode, TwGridSelectionFree)
Call PropBag.WriteProperty("ScrollBars", TickerGrid.ScrollBars, TwGridScrollBarBoth)
Call PropBag.WriteProperty("RowSizingMode", TickerGrid.RowSizingMode, TwGridRowSizeAll)
Call PropBag.WriteProperty("Rows", TickerGrid.Rows, GridRowsInitial)
Call PropBag.WriteProperty("RowHeightMin", TickerGrid.RowHeightMin, 0)
Call PropBag.WriteProperty("RowBackColorOdd", TickerGrid.RowBackColorOdd, 0)
Call PropBag.WriteProperty("RowBackColorEven", TickerGrid.RowBackColorEven, 0)
Call PropBag.WriteProperty("HighLight", TickerGrid.HighLight, TwGridHighlightWithFocus)
Call PropBag.WriteProperty("GridLinesFixed", TickerGrid.GridLinesFixed, TwGridGridFlat)
Call PropBag.WriteProperty("GridLines", TickerGrid.GridLines, TwGridGridNone)
Call PropBag.WriteProperty("GridLineWidth", TickerGrid.GridLineWidth, 1)
Call PropBag.WriteProperty("GridColorFixed", TickerGrid.GridColorFixed, 12632256)
Call PropBag.WriteProperty("GridColor", TickerGrid.GridColor, -2147483631)
Call PropBag.WriteProperty("ForeColorSel", TickerGrid.ForeColorSel, -2147483634)
Call PropBag.WriteProperty("ForeColorFixed", TickerGrid.ForeColorFixed, -2147483630)
Call PropBag.WriteProperty("ForeColor", TickerGrid.ForeColor, &H80000008)
Call PropBag.WriteProperty("FontFixed", TickerGrid.FontFixed)
Call PropBag.WriteProperty("Font", TickerGrid.Font)
Call PropBag.WriteProperty("FocusRect", TickerGrid.FocusRect, TwGridFocusBroken)
Call PropBag.WriteProperty("FixedRows", TickerGrid.FixedRows, 1)
Call PropBag.WriteProperty("FixedCols", TickerGrid.FixedCols, 1)
Call PropBag.WriteProperty("FillStyle", TickerGrid.FillStyle, TwGridFillRepeat)
Call PropBag.WriteProperty("Cols", TickerGrid.Cols, 2)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

Dim lTicker As IMarketDataSource
Set lTicker = Data

removeTickerFromConfig lTicker

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IErrorListener Interface Members
'@================================================================================

Private Sub IErrorListener_Notify(ev As ErrorEventData)
Const ProcName As String = "IErrorListener_Notify"
On Error GoTo Err

RaiseEvent Error(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' IPriceChangeListener Interface Members
'@================================================================================

Private Sub IPriceChangeListener_Change(ev As PriceChangeEventData)
Const ProcName As String = "IPriceChangeListener_Change"
On Error GoTo Err

Dim lDataSource As IMarketDataSource
Set lDataSource = ev.Source

TickerGrid.BeginCellEdit getTickerGridRow(lDataSource), mColumnMap(TickerGridColumns.Change)
TickerGrid.Text = ev.PriceChange.ChangeString
If ev.PriceChange.Change >= 0 Then
    TickerGrid.CellBackColor = mPositiveChangeBackColor
    TickerGrid.CellForeColor = mPositiveChangeForeColor
Else
    TickerGrid.CellBackColor = mNegativeChangeBackColor
    TickerGrid.CellForeColor = mNegativeChangeForeColor
End If
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit getTickerGridRow(lDataSource), mColumnMap(TickerGridColumns.ChangePercent)
TickerGrid.Text = Format(ev.PriceChange.ChangePercent, "0.0")
If ev.PriceChange.ChangePercent >= 0 Then
    TickerGrid.CellBackColor = mPositiveChangeBackColor
    TickerGrid.CellForeColor = mPositiveChangeForeColor
Else
    TickerGrid.CellBackColor = mNegativeChangeBackColor
    TickerGrid.CellForeColor = mNegativeChangeForeColor
End If
TickerGrid.EndCellEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IQuoteListener Interface Members
'@================================================================================

Private Sub IQuoteListener_Ask(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_Ask"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Ask)
displaySize ev, mColumnMap(TickerGridColumns.AskSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_bid(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_bid"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Bid)
displaySize ev, mColumnMap(TickerGridColumns.BidSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_high(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_high"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.HighPrice)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_Low(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_Low"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.LowPrice)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_openInterest(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_openInterest"
On Error GoTo Err

displaySize ev, mColumnMap(TickerGridColumns.OpenInterest)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_previousClose(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_previousClose"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.ClosePrice)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_sessionOpen(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_sessionOpen"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.OpenPrice)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_trade(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_trade"
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Trade)
displaySize ev, mColumnMap(TickerGridColumns.TradeSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IQuoteListener_volume(ev As QuoteEventData)
Const ProcName As String = "IQuoteListener_volume"
On Error GoTo Err

displaySize ev, mColumnMap(TickerGridColumns.Volume)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

processTickerState ev.Source

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TickerGrid_Click()
Const ProcName As String = "TickerGrid_Click"
On Error GoTo Err


Dim lRow As Long
lRow = TickerGrid.Row

Dim lRowSel As Long
lRowSel = TickerGrid.RowSel

Dim lCol As Long
lCol = TickerGrid.col

Dim lColSel As Long
lColSel = TickerGrid.ColSel

mSelectedTickers.BeginChange
If lCol = 1 And lColSel = TickerGrid.Cols - 1 And _
    lRow = 1 And lRowSel = TickerGrid.Rows - 1 _
Then
    ' the user has clicked on the top left cell so select all rows
    ' regardless of whether ctrl is down
    
    DeselectSelectedTickers
    SelectAllTickers
ElseIf mShiftDown Then
    DeselectSelectedTickers
    
    Dim i As Long
    For i = TickerGrid.Row To TickerGrid.RowSel Step IIf(TickerGrid.Row <= TickerGrid.RowSel, 1, -1)
        SelectTicker i
    Next
ElseIf mControlDown Then
    toggleRowSelection lRow
Else
    DeselectSelectedTickers
    SelectTicker lRow
End If
mSelectedTickers.EndChange

RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_ColMoved( _
                ByVal pFromCol As Long, _
                ByVal pToCol As Long)
Const ProcName As String = "TickerGrid_ColMoved"
On Error GoTo Err

adjustMovedColumn pFromCol, pToCol

RaiseEvent ColMoved(pFromCol, pToCol)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean)
Const ProcName As String = "TickerGrid_ColMoving"
On Error GoTo Err

RaiseEvent ColMoving(fromCol, toCol, Cancel)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_DblClick()
Const ProcName As String = "TickerGrid_DblClick"
On Error GoTo Err

RaiseEvent DblClick

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TickerGrid_KeyDown"
On Error GoTo Err

RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_KeyPress(KeyAscii As Integer)
Const ProcName As String = "TickerGrid_KeyPress"
On Error GoTo Err

RaiseEvent KeyPress(KeyAscii)

If isAlphaNumeric(KeyAscii) Then processAlphaNumericKey KeyAscii

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TickerGrid_KeyUp"
On Error GoTo Err

RaiseEvent KeyUp(KeyCode, Shift)

Select Case KeyCode
'Case vbKeyDelete
'   Let the application handle this as it may want to Finish them as well
'   as stopping them
'    StopSelectedTickers
Case vbKeyInsert
    insertBlankRow TickerGrid.Row
Case vbKeyUp
    moveSelectedRows -1
Case vbKeyDown
    moveSelectedRows 1
Case vbKeyEscape
    stopEnteringTickerSymbol
Case vbKeyBack
    truncateTickerSymbol
Case vbKeyReturn
    notifyTickerSymbolEntry
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "TickerGrid_MouseDown"
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "TickerGrid_MouseMove"
On Error GoTo Err

RaiseEvent MouseMove(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "TickerGrid_MouseUp"
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseUp(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_RowMoved( _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long)
Const ProcName As String = "TickerGrid_RowMoved"
On Error GoTo Err

adjustMovedRow pFromRow, pToRow
RaiseEvent RowMoved(pFromRow, pToRow)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_RowMoving( _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long, _
                Cancel As Boolean)
Const ProcName As String = "TickerGrid_RowMoving"
On Error GoTo Err

RaiseEvent RowMoving(pFromRow, pToRow, Cancel)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TickerGrid_SelectionChanged(ByVal pRow1 As Long, ByVal pCol1 As Long, ByVal pRow2 As Long, ByVal pCol2 As Long)
Const ProcName As String = "TickerGrid_SelectionChanged"
On Error GoTo Err

stopEnteringTickerSymbol
RaiseEvent SelectionChanged(pRow1, pCol1, pRow2, pCol2)
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mRefreshPriceChangeTC Event Handlers
'@================================================================================

Private Sub mRefreshPriceChangeTC_Completed(ev As TWUtilities40.TaskCompletionEventData)
Const ProcName As String = "mRefreshPriceChangeTC_Completed"
On Error GoTo Err

Set mRefreshPriceChangeTC = Nothing
If mRefreshQuotesTC Is Nothing Then Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mRefreshQuotesTC Event Handlers
'@================================================================================

Private Sub mRefreshQuotesTC_Completed(ev As TWUtilities40.TaskCompletionEventData)
Const ProcName As String = "mRefreshQuotesTC_Completed"
On Error GoTo Err

Set mRefreshQuotesTC = Nothing
If mRefreshPriceChangeTC Is Nothing Then Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mSelectedTickers Event Handlers
'@================================================================================

Private Sub mSelectedTickers_SelectionChanged()
Const ProcName As String = "mSelectedTickers_SelectionChanged"
On Error GoTo Err

RaiseEvent TickerSelectionChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

TickerGrid.Theme = mTheme
IncreasedValueColor = mTheme.IncreasedValueColor
DecreasedValueColor = mTheme.DecreasedValueColor
NegativeChangeBackColor = mTheme.NegativeChangeBackColor
NegativeChangeForeColor = mTheme.NegativeChangeForeColor
PositiveChangeBackColor = mTheme.PositiveChangeBackColor
PositiveChangeForeColor = mTheme.PositiveChangeForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Get SelectedTickers() As SelectedTickers
Attribute SelectedTickers.VB_MemberFlags = "400"
Const ProcName As String = "SelectedTickers"
On Error GoTo Err

Set SelectedTickers = mSelectedTickers

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ScrollBars() As TWControls40.ScrollBarsSettings
Const ProcName As String = "ScrollBars"
On Error GoTo Err

ScrollBars = TickerGrid.ScrollBars

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As TWControls40.ScrollBarsSettings)
Attribute ScrollBars.VB_Description = "Specifies whether scroll bars are to be provided."
Attribute ScrollBars.VB_ProcData.VB_Invoke_PropertyPut = ";Behavior"
Const ProcName As String = "ScrollBars"
On Error GoTo Err

TickerGrid.ScrollBars = New_ScrollBars
PropertyChanged "ScrollBars"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowSizingMode() As RowSizingSettings
Attribute RowSizingMode.VB_Description = "Specifies whether resizing a row affects only that row or all rows."
Attribute RowSizingMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "RowSizingMode"
On Error GoTo Err

RowSizingMode = TickerGrid.RowSizingMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowSizingMode(ByVal New_RowSizingMode As RowSizingSettings)
Const ProcName As String = "RowSizingMode"
On Error GoTo Err

TickerGrid.RowSizingMode = New_RowSizingMode
PropertyChanged "RowSizingMode"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Specifies the initial number of rows (bear in mind that the header consumes one row)."
Attribute Rows.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "Rows"
On Error GoTo Err

Rows = TickerGrid.Rows

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Rows(ByVal New_Rows As Long)
Const ProcName As String = "Rows"
On Error GoTo Err

TickerGrid.Rows = New_Rows
PropertyChanged "Rows"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Specifies the minimum height to which a row can be resized by the user."
Attribute RowHeightMin.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "RowHeightMin"
On Error GoTo Err

    RowHeightMin = TickerGrid.RowHeightMin

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
Const ProcName As String = "RowHeightMin"
On Error GoTo Err

    TickerGrid.RowHeightMin = New_RowHeightMin
    PropertyChanged "RowHeightMin"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowForeColorOdd() As OLE_COLOR
Const ProcName As String = "RowForeColorOdd"
On Error GoTo Err

    RowForeColorOdd = TickerGrid.RowForeColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowForeColorOdd(ByVal New_RowForeColorOdd As OLE_COLOR)
Const ProcName As String = "RowForeColorOdd"
On Error GoTo Err

    TickerGrid.RowForeColorOdd = New_RowForeColorOdd
    PropertyChanged "RowForeColorOdd"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowForeColorEven() As OLE_COLOR
Const ProcName As String = "RowForeColorEven"
On Error GoTo Err

    RowForeColorEven = TickerGrid.RowForeColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowForeColorEven(ByVal New_RowForeColorEven As OLE_COLOR)
Const ProcName As String = "RowForeColorEven"
On Error GoTo Err

    TickerGrid.RowForeColorEven = New_RowForeColorEven
    PropertyChanged "RowForeColorEven"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
Attribute RowBackColorOdd.VB_Description = "Specifies the background color for odd-numbered rows."
Attribute RowBackColorOdd.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

    RowBackColorOdd = TickerGrid.RowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorOdd(ByVal New_RowBackColorOdd As OLE_COLOR)
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

    TickerGrid.RowBackColorOdd = New_RowBackColorOdd
    PropertyChanged "RowBackColorOdd"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
Attribute RowBackColorEven.VB_Description = "Specifies the background color for even-numbered rows."
Attribute RowBackColorEven.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute RowBackColorEven.VB_MemberFlags = "200"
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

    RowBackColorEven = TickerGrid.RowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorEven(ByVal New_RowBackColorEven As OLE_COLOR)
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

    TickerGrid.RowBackColorEven = New_RowBackColorEven
    PropertyChanged "RowBackColorEven"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
Const ProcName As String = "Redraw"
On Error GoTo Err

Redraw = TickerGrid.Redraw

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Redraw(ByVal Value As Boolean)
Const ProcName As String = "Redraw"
On Error GoTo Err

TickerGrid.Redraw = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PositiveChangeBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "PositiveChangeBackColor"
On Error GoTo Err

mPositiveChangeBackColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPositiveChangeBackColor, mPositiveChangeBackColor
RefreshPriceChange
PropertyChanged "PositiveChangeBackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositiveChangeBackColor() As OLE_COLOR
Attribute PositiveChangeBackColor.VB_Description = "Specifies the background color for price change cells when the price has increased."
Attribute PositiveChangeBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "PositiveChangeBackColor"
On Error GoTo Err

PositiveChangeBackColor = mPositiveChangeBackColor
PropertyChanged "PositiveChangeBackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PositiveChangeForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "PositiveChangeForeColor"
On Error GoTo Err

mPositiveChangeForeColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPositiveChangeForeColor, mPositiveChangeForeColor
RefreshPriceChange
PropertyChanged "PositiveChangeForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PositiveChangeForeColor() As OLE_COLOR
Attribute PositiveChangeForeColor.VB_Description = "Specifies the foreground color for price change cells when the price has increased."
Attribute PositiveChangeForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "PositiveChangeForeColor"
On Error GoTo Err

PositiveChangeForeColor = mPositiveChangeForeColor
PropertyChanged "PositiveChangeForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let NegativeChangeBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "NegativeChangeBackColor"
On Error GoTo Err

mNegativeChangeBackColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingNegativeChangeBackColor, mNegativeChangeBackColor
RefreshPriceChange
PropertyChanged "NegativeChangeBackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NegativeChangeBackColor() As OLE_COLOR
Attribute NegativeChangeBackColor.VB_Description = "Specifies the background color for price change cells when the price has decreased."
Attribute NegativeChangeBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "NegativeChangeBackColor"
On Error GoTo Err

NegativeChangeBackColor = mNegativeChangeBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let NegativeChangeForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "NegativeChangeForeColor"
On Error GoTo Err

mNegativeChangeForeColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingNegativeChangeForeColor, mNegativeChangeForeColor
RefreshPriceChange
PropertyChanged "NegativeChangeForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NegativeChangeForeColor() As OLE_COLOR
Attribute NegativeChangeForeColor.VB_Description = "Specifies the foreground color for price change cells when the price has decreased."
Attribute NegativeChangeForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "NegativeChangeForeColor"
On Error GoTo Err

NegativeChangeForeColor = mNegativeChangeForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IncreasedValueColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "IncreasedValueColor"
On Error GoTo Err

mIncreasedValueColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingIncreasedValueColor, mIncreasedValueColor
RefreshQuotes
PropertyChanged "IncreasedValueColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncreasedValueColor() As OLE_COLOR
Attribute IncreasedValueColor.VB_Description = "Specifies the foreground color for price cells that have increased in value."
Attribute IncreasedValueColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "IncreasedValueColor"
On Error GoTo Err

IncreasedValueColor = mIncreasedValueColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HighlightPriceChanges(ByVal Value As Boolean)
Const ProcName As String = "HighlightPriceChanges"
On Error GoTo Err

mHighlightPriceChanges = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHighlightPriceChanges, mHighlightPriceChanges
RefreshPriceChange
PropertyChanged "HighlightPriceChanges"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HighlightPriceChanges() As Boolean
Attribute HighlightPriceChanges.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "HighlightPriceChanges"
On Error GoTo Err

HighlightPriceChanges = mHighlightPriceChanges

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
hWnd = UserControl.hWnd
End Property

Public Property Get GridLineWidth() As Long
Attribute GridLineWidth.VB_Description = "Specifies the thickness of the grid lines."
Attribute GridLineWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "GridLineWidth"
On Error GoTo Err

    GridLineWidth = TickerGrid.GridLineWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Long)
Const ProcName As String = "GridLineWidth"
On Error GoTo Err

    TickerGrid.GridLineWidth = New_GridLineWidth
    PropertyChanged "GridLineWidth"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_Description = "Specifies the color of the header grid lines."
Attribute GridColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "GridColorFixed"
On Error GoTo Err

    GridColorFixed = TickerGrid.GridColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
Const ProcName As String = "GridColorFixed"
On Error GoTo Err

    TickerGrid.GridColorFixed = New_GridColorFixed
    PropertyChanged "GridColorFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Specifies the color of the grid lines."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "GridColor"
On Error GoTo Err

    GridColor = TickerGrid.GridColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
Const ProcName As String = "GridColor"
On Error GoTo Err

    TickerGrid.GridColor = New_GridColor
    PropertyChanged "GridColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Specifies the foreground color for header cells."
Attribute ForeColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "ForeColorFixed"
On Error GoTo Err

    ForeColorFixed = TickerGrid.ForeColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
Const ProcName As String = "ForeColorFixed"
On Error GoTo Err

    TickerGrid.ForeColorFixed = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Specifies the foreground color for non-header cells."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
Const ProcName As String = "foreColor"
On Error GoTo Err

    ForeColor = TickerGrid.ForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

    TickerGrid.ForeColor = New_ForeColor
    PropertyChanged "ForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FontFixed() As StdFont
Attribute FontFixed.VB_Description = "Specifies the font to be used for header cells."
Attribute FontFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "FontFixed"
On Error GoTo Err

Set FontFixed = TickerGrid.FontFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let FontFixed(ByVal Value As StdFont)
Const ProcName As String = "FontFixed"
On Error GoTo Err

Set TickerGrid.FontFixed = Value
PropertyChanged "FontFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set FontFixed(ByVal Value As StdFont)
Const ProcName As String = "FontFixed"
On Error GoTo Err

TickerGrid.FontFixed = Value
PropertyChanged "FontFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Specifies the font to be used for non-header cells."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
Const ProcName As String = "Font"
On Error GoTo Err

Set Font = TickerGrid.Font

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set Font(ByVal Value As StdFont)
Const ProcName As String = "Font"
On Error GoTo Err

Set TickerGrid.Font = Value
calcAverageCharacterWidths Value
setColumnWidths
PropertyChanged "FontFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Font(ByVal Value As StdFont)
Const ProcName As String = "Font"
On Error GoTo Err

Set TickerGrid.Font = Value
calcAverageCharacterWidths Value
setColumnWidths
PropertyChanged "FontFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DecreasedValueColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "DecreasedValueColor"
On Error GoTo Err

mDecreasedValueColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDecreasedValueColor, mDecreasedValueColor
RefreshQuotes
PropertyChanged "DecreasedValueColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DecreasedValueColor() As OLE_COLOR
Attribute DecreasedValueColor.VB_Description = "Specifies the foreground color for price cells that have decreased in value."
Attribute DecreasedValueColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "DecreasedValueColor"
On Error GoTo Err

DecreasedValueColor = mDecreasedValueColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If mConfig Is Value Then Exit Property
If Not mConfig Is Nothing Then mConfig.Remove
If Value Is Nothing Then Exit Property

Set mConfig = Value
Set mTickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)

mConfig.SetSetting ConfigSettingPositiveChangeBackColor, mPositiveChangeBackColor
mConfig.SetSetting ConfigSettingPositiveChangeForeColor, mPositiveChangeForeColor
mConfig.SetSetting ConfigSettingNegativeChangeBackColor, mNegativeChangeBackColor
mConfig.SetSetting ConfigSettingNegativeChangeForeColor, mNegativeChangeForeColor
mConfig.SetSetting ConfigSettingIncreasedValueColor, mIncreasedValueColor
mConfig.SetSetting ConfigSettingDecreasedValueColor, mDecreasedValueColor

storeColumnMap

TickerGrid.ConfigurationSection = mConfig.AddPrivateConfigurationSection(ConfigSectionGrid)

Dim lTicker As IMarketDataSource
For Each lTicker In mTickers
    storeTickerSettings lTicker
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BorderStyle() As TWUtilities40.BorderStyleSettings
Const ProcName As String = "BorderStyle"
On Error GoTo Err

BorderStyle = TickerGrid.BorderStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BorderStyle(ByVal Value As TWUtilities40.BorderStyleSettings)
Const ProcName As String = "BorderStyle"
On Error GoTo Err

TickerGrid.BorderStyle = Value
PropertyChanged "BorderStyle"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Specifies the background color of the fixed cells (ie row and column headers)."
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "BackColorFixed"
On Error GoTo Err

    BackColorFixed = TickerGrid.BackColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
Const ProcName As String = "BackColorFixed"
On Error GoTo Err

    TickerGrid.BackColorFixed = New_BackColorFixed
    PropertyChanged "BackColorFixed"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Specifies the color of the area behind the rows and columns."
Attribute BackColorBkg.VB_ProcData.VB_Invoke_Property = ";Appearance"
Const ProcName As String = "BackColorBkg"
On Error GoTo Err

    BackColorBkg = TickerGrid.BackColorBkg

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
Const ProcName As String = "BackColorBkg"
On Error GoTo Err

TickerGrid.BackColorBkg = New_BackColorBkg
PropertyChanged "BackColorBkg"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
Attribute BackColor.VB_MemberFlags = "400"
Const ProcName As String = "backColor"
On Error GoTo Err

    BackColor = TickerGrid.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Const ProcName As String = "backColor"
On Error GoTo Err

    TickerGrid.BackColor = New_BackColor
    PropertyChanged "BackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AllowUserResizing() As AllowUserResizeSettings
Attribute AllowUserResizing.VB_Description = "Specifies whethe the user is allowed to change the size of columns and/or rows."
Attribute AllowUserResizing.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "AllowUserResizing"
On Error GoTo Err

    AllowUserResizing = TickerGrid.AllowUserResizing

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As AllowUserResizeSettings)
Const ProcName As String = "AllowUserResizing"
On Error GoTo Err

    TickerGrid.AllowUserResizing = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get AllowUserReordering() As AllowUserReorderSettings
Attribute AllowUserReordering.VB_Description = "Specifies whether the user is allowed to change the order of columns and/or rows."
Attribute AllowUserReordering.VB_ProcData.VB_Invoke_Property = ";Behavior"
Const ProcName As String = "AllowUserReordering"
On Error GoTo Err

    AllowUserReordering = TickerGrid.AllowUserReordering

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let AllowUserReordering(ByVal New_AllowUserReordering As AllowUserReorderSettings)
Const ProcName As String = "AllowUserReordering"
On Error GoTo Err

    TickerGrid.AllowUserReordering = New_AllowUserReordering
    PropertyChanged "AllowUserReordering"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddTickerFromDataSource(ByVal pDataSource As IMarketDataSource, Optional ByVal pGridRow As Long) As IMarketDataSource
Const ProcName As String = "AddTickerFromDataSource"
On Error GoTo Err

addTickerToGrid pDataSource, pGridRow
processTickerState pDataSource

If Not pDataSource.IsMarketDataRequested Then pDataSource.StartMarketData

Set AddTickerFromDataSource = pDataSource

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub DeselectSelectedTickers()
Const ProcName As String = "DeselectSelectedTickers"
On Error GoTo Err

Dim lIndex As Long
Dim i As Long
Dim lTicker As IMarketDataSource

For Each lTicker In mTickers
    lIndex = getTickerIndex(lTicker)
    If isTickerSelected(lIndex) Then
        toggleRowHighlight getTickerGridRowFromIndex(lIndex)
    End If
Next

mSelectedTickers.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DeselectTicker( _
                ByVal pIndex As Long)
Const ProcName As String = "deselectTicker"
On Error GoTo Err

deselectATickerByRow pIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub ExtendSelection(ByVal pRow As Long, ByVal pCol As Long)
'    TickerGrid.ExtendSelection row, col
'End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

gLogger.Log "Finishing TickerGrid", ProcName, ModuleName, LogLevelDetail

Dim lTicker As IMarketDataSource
For Each lTicker In mTickers
    stopListeningToTicker lTicker
Next

Set mTickers = Nothing
TickerGrid.Clear
ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
mSelectedTickers.Clear

Exit Sub
Err:
'ignore any errors
End Sub

Public Function GetColFromX(ByVal X As Long) As Long
Const ProcName As String = "GetColFromX"
On Error GoTo Err

    GetColFromX = TickerGrid.GetColFromX(X)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetRowFromY(ByVal Y As Long) As Long
Const ProcName As String = "GetRowFromY"
On Error GoTo Err

    GetRowFromY = TickerGrid.GetRowFromY(Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub InvertCellColors()
Const ProcName As String = "InvertCellColors"
On Error GoTo Err

TickerGrid.InvertCellColors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pMarketDataManager As IMarketDataManager, _
                Optional ByVal pConfig As ConfigurationSection)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pMarketDataManager Is Nothing, "pMarketDataManager must not be Nothing"
Set mMarketDataManager = pMarketDataManager

setupDefaultTickerGridColumns
setupDefaultTickerGridHeaders

If Not pConfig Is Nothing Then LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Assert Not mMarketDataManager Is Nothing, "TickerGrid has not been initialised"
AssertArgument Not pConfig Is Nothing, "pConfig cannot be Nothing"

Set mConfig = pConfig
If mConfig Is Nothing Then Exit Sub

Set mTickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)

TickerGrid.Redraw = False
TickerGrid.LoadFromConfig mConfig.AddPrivateConfigurationSection(ConfigSectionGrid)
TickerGrid.Redraw = True

loadColumnMap

' adjust columns to take account of column map
setupDefaultTickerGridColumns
setupDefaultTickerGridHeaders

If mConfig.GetSetting(ConfigSettingPositiveChangeBackColor) <> "" Then mPositiveChangeBackColor = mConfig.GetSetting(ConfigSettingPositiveChangeBackColor)
If mConfig.GetSetting(ConfigSettingPositiveChangeForeColor) <> "" Then mPositiveChangeForeColor = mConfig.GetSetting(ConfigSettingPositiveChangeForeColor)
If mConfig.GetSetting(ConfigSettingNegativeChangeBackColor) <> "" Then mNegativeChangeBackColor = mConfig.GetSetting(ConfigSettingNegativeChangeBackColor)
If mConfig.GetSetting(ConfigSettingNegativeChangeForeColor) <> "" Then mNegativeChangeForeColor = mConfig.GetSetting(ConfigSettingNegativeChangeForeColor)
If mConfig.GetSetting(ConfigSettingIncreasedValueColor) <> "" Then mIncreasedValueColor = mConfig.GetSetting(ConfigSettingIncreasedValueColor)
If mConfig.GetSetting(ConfigSettingHighlightPriceChanges) <> "" Then mHighlightPriceChanges = mConfig.GetSetting(ConfigSettingHighlightPriceChanges)
If mConfig.GetSetting(ConfigSettingDecreasedValueColor) <> "" Then mDecreasedValueColor = mConfig.GetSetting(ConfigSettingDecreasedValueColor)

mIteratingTickersConfig = True
Dim lTickerConfig As ConfigurationSection
For Each lTickerConfig In mTickersConfigSection
    gLogger.Log "Recovering Ticker " & lTickerConfig.InstanceQualifier, ProcName, ModuleName, LogLevelNormal
    Dim lTicker As IMarketDataSource
    Set lTicker = mMarketDataManager.GetMarketDataSource(lTickerConfig.InstanceQualifier)
    addTickerToGrid lTicker, CLng(lTickerConfig.GetSetting(ConfigSettingRowIndex, "0"))
    processTickerState lTicker
    
    If Not lTicker.IsMarketDataRequested Then lTicker.StartMarketData
    
    gLogger.Log "Started Ticker " & gGetContractFromContractFuture(lTicker.ContractFuture).Specifier.ToString, ProcName, ModuleName, LogLevelNormal
Next
mIteratingTickersConfig = False

gLogger.Log "Ticker grid loaded from pConfig", ProcName, ModuleName, LogLevelNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MoveColumn(ByVal pFromCol As Long, ByVal pToCol As Long)
Const ProcName As String = "MoveColumn"
On Error GoTo Err

TickerGrid.MoveColumn pFromCol, pToCol
adjustMovedColumn pFromCol, pToCol

RaiseEvent ColMoved(pFromCol, pToCol)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MoveRow(ByVal pFromRow As Long, ByVal pToRow As Long)
Const ProcName As String = "MoveRow"
On Error GoTo Err

TickerGrid.MoveRow pFromRow, pToRow
adjustMovedRow pFromRow, pToRow

RaiseEvent RowMoved(pFromRow, pToRow)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToCell(pRow As Long, pCol As Long)
Const ProcName As String = "ScrollToCell"
On Error GoTo Err

TickerGrid.ScrollToCell pRow, pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToCol(pCol As Long)
Const ProcName As String = "ScrollToCol"
On Error GoTo Err

TickerGrid.ScrollToCol pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToRow(pRow As Long)
Const ProcName As String = "ScrollToRow"
On Error GoTo Err

TickerGrid.ScrollToRow pRow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectAllTickers()
Const ProcName As String = "SelectAllTickers"
On Error GoTo Err

Dim lTicker As IMarketDataSource

mSelectedTickers.BeginChange
For Each lTicker In mTickers
    selectATickerByIndex getTickerIndex(lTicker)
Next
mSelectedTickers.EndChange

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectTicker( _
                ByVal pRow As Long)
Const ProcName As String = "SelectTicker"
On Error GoTo Err

selectATickerByRow pRow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetCellAlignment(ByVal pRow As Long, ByVal pCol As Long, pAlign As TWControls40.AlignmentSettings)
Const ProcName As String = "setCellAlignment"
On Error GoTo Err

TickerGrid.SetCellAlignment pRow, pCol, pAlign

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartTickerFromContract(ByVal pContract As IContract, Optional ByVal pGridRow As Long) As IMarketDataSource
Const ProcName As String = "StartTickerFromContract"
On Error GoTo Err

AssertArgument Not IsContractExpired(pContract), "Contract has expired"

Set StartTickerFromContract = StartTickerFromContractFuture(CreateFuture(pContract), pGridRow)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function StartTickerFromContractFuture(ByVal pContractFuture As IFuture, Optional ByVal pGridRow As Long) As IMarketDataSource
Const ProcName As String = "StartTickerFromContractFuture"
On Error GoTo Err

Dim lTicker As IMarketDataSource
Set lTicker = mMarketDataManager.CreateMarketDataSource(pContractFuture, True)

addTickerToGrid lTicker, pGridRow
processTickerState lTicker

lTicker.StartMarketData

Set StartTickerFromContractFuture = lTicker

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub StopAllTickers()
Const ProcName As String = "StopAllTickers"
On Error GoTo Err

Dim i As Long

TickerGrid.Redraw = False

' do this in reverse order - most efficient when all tickers are being stopped
For i = TickerGrid.Rows - 1 To 1 Step -1
    If isRowOccupied(i) Then
        stopTicker getTickerFromGridRow(i)
    End If
Next

TickerGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub StopSelectedTickers()
Const ProcName As String = "StopSelectedTickers"
On Error GoTo Err

Dim lTicker As IMarketDataSource

TickerGrid.Redraw = False

Dim lSelectedTickers As SelectedTickers
Set lSelectedTickers = mSelectedTickers
Set mSelectedTickers = New SelectedTickers

For Each lTicker In lSelectedTickers
    stopTicker lTicker
Next

TickerGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addTickerToGrid(ByVal pTicker As IMarketDataSource, ByVal pGridRow As Long)
Const ProcName As String = "addTickerToGrid"
On Error GoTo Err

listenToTicker pTicker

Dim lIndex As Long
lIndex = addTickerToTickerTable(pTicker)

Dim lRow As Long
If pGridRow > 0 Then
    If isRowOccupied(pGridRow) Then insertBlankRow pGridRow
'    If isRowOccupied(pGridRow) Then
'        If getTickerFromGridRow(pGridRow).State <> MarketDataSourceStateError Then insertBlankRow pGridRow
'    End If
    lRow = pGridRow
Else
    lRow = allocateRow
End If

setTickerGridRowFromIndex lIndex, lRow
setTickerIndexForRow lRow, lIndex

gLogger.Log "Added Ticker to grid " & pTicker.Key, ProcName, ModuleName, LogLevelNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addTickerToTickerTable(pTicker) As Long
Const ProcName As String = "addTickerToTickerTable"
On Error GoTo Err

Dim lIndex As Long

lIndex = getTickerIndex(pTicker)

Do While lIndex > UBound(mTickerTable)
    ReDim Preserve mTickerTable((UBound(mTickerTable) + 1) * TickerTableEntriesGrowthFactor - 1) As TickerTableEntry
Loop

Set mTickerTable(lIndex).DataSource = pTicker
mTickers.Add pTicker, pTicker.Key

addTickerToTickerTable = lIndex
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub adjustMovedColumn(ByVal pFromCol As Long, ByVal pToCol As Long)
Const ProcName As String = "adjustMovedColumn"
On Error GoTo Err

Dim i As Long

For i = pFromCol To pToCol Step IIf(pFromCol <= pToCol, 1, -1)
    mColumnMap(TickerGrid.ColData(i)) = i
Next

storeColumnMap

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub adjustMovedRow(ByVal pFromRow As Long, ByVal pToRow As Long)
Const ProcName As String = "adjustMovedRow"
On Error GoTo Err

Dim i As Long
Dim lTicker As IMarketDataSource

For i = pFromRow To pToRow Step IIf(pFromRow <= pToRow, 1, -1)
    Set lTicker = getTickerFromGridRow(i)
    If Not lTicker Is Nothing Then setTickerGridRow lTicker, i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateRow() As Long
Const ProcName As String = "allocateRow"
On Error GoTo Err

Dim i As Long
For i = 1 To TickerGrid.Rows - 1
    If isRowOccupied(i) Then
'        If getTickerFromGridRow(i).State = MarketDataSourceStateError Then
'            allocateRow = i
'            Exit For
'        End If
    Else
        allocateRow = i
        Exit For
    End If
Next

If allocateRow > TickerGrid.Rows - GridRowsIncrement Then
    TickerGrid.Rows = TickerGrid.Rows + GridRowsIncrement
End If


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calcAverageCharacterWidths( _
                ByVal afont As StdFont)
Const ProcName As String = "calcAverageCharacterWidths"
On Error GoTo Err

mLetterWidth = getAverageCharacterWidth("ABCDEFGH IJKLMNOP QRST UVWX YZ", afont)
mDigitWidth = getAverageCharacterWidth(".0123456789", afont)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function cloneTickers() As EnumerableCollection
Const ProcName As String = "cloneTickers"
On Error GoTo Err

Dim lTickers As New EnumerableCollection
Dim lTicker As IMarketDataSource

For Each lTicker In mTickers
    lTickers.Add lTicker, lTicker.Key
Next

Set cloneTickers = lTickers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub deselectATickerByIndex( _
                ByVal pIndex As Long)
Const ProcName As String = "deselectATickerByRow"
On Error GoTo Err

If isTickerSelected(pIndex) Then
    mSelectedTickers.Remove getTickerFromIndex(pIndex)
    toggleRowHighlight getTickerGridRowFromIndex(pIndex)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deselectATickerByRow( _
                ByVal pRow As Long)
Const ProcName As String = "deselectATickerByRow"
On Error GoTo Err

deselectATickerByIndex getTickerIndexFromGridRow(pRow)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displayPrice( _
                ev As QuoteEventData, _
                ByVal pCol As Long)
Const ProcName As String = "displayPrice"
On Error GoTo Err

Dim lTicker As IMarketDataSource

Set lTicker = ev.Source
TickerGrid.BeginCellEdit getTickerGridRow(lTicker), pCol
TickerGrid.Text = GetFormattedPriceFromQuoteEvent(ev)

With ev.Quote
    If .PriceChange = ValueChangeNone Or (Not mHighlightPriceChanges) Then
        If .RecentPriceChange = ValueChangeUp Then
            TickerGrid.CellForeColor = mIncreasedValueColor
        ElseIf .RecentPriceChange = ValueChangeDown Then
            TickerGrid.CellForeColor = mDecreasedValueColor
        Else
            TickerGrid.CellForeColor = 0
        End If
        
        TickerGrid.CellBackColor = 0
    Else
        TickerGrid.CellBackColor = 0    ' reset backcolor to default
        TickerGrid.CellForeColor = 0    ' TickerGrid.CellBackColor
        If .PriceChange = ValueChangeUp Then
            TickerGrid.CellBackColor = mIncreasedValueColor
        ElseIf .PriceChange = ValueChangeDown Then
            TickerGrid.CellBackColor = mDecreasedValueColor
        End If
    End If
End With
TickerGrid.EndCellEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub displaySize( _
                ev As QuoteEventData, _
                ByVal pCol As Long)
Const ProcName As String = "displaySize"
On Error GoTo Err

Dim lTicker As IMarketDataSource

Set lTicker = ev.Source
With ev.Quote
    TickerGrid.BeginCellEdit getTickerGridRow(lTicker), pCol
    TickerGrid.Text = .Size
    
    If .SizeChange = ValueChangeNone Then
        If .RecentSizeChange = ValueChangeUp Then
            TickerGrid.CellForeColor = mIncreasedValueColor
        ElseIf .RecentSizeChange = ValueChangeDown Then
            TickerGrid.CellForeColor = mDecreasedValueColor
        Else
            TickerGrid.CellForeColor = 0
        End If
    Else
        If .SizeChange = ValueChangeUp Then
            TickerGrid.CellForeColor = mIncreasedValueColor
        ElseIf .SizeChange = ValueChangeDown Then
            TickerGrid.CellForeColor = mDecreasedValueColor
        End If
    End If
    
    TickerGrid.EndCellEdit
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getAverageCharacterWidth( _
                ByVal widthString As String, _
                ByVal pFont As StdFont) As Long
Const ProcName As String = "getAverageCharacterWidth"
On Error GoTo Err

Set FontPicture.Font = pFont
getAverageCharacterWidth = FontPicture.TextWidth(widthString) / Len(widthString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFieldsHaveBeenSet(ByVal pIndex As Long) As Boolean
getFieldsHaveBeenSet = mTickerTable(pIndex).FieldsHaveBeenSet
End Function

Private Function getTickerFromGridRow(ByVal pRow As Long) As IMarketDataSource
Set getTickerFromGridRow = getTickerFromIndex(getTickerIndexFromGridRow(pRow))
End Function

Private Function getTickerFromIndex(ByVal pIndex As Long) As IMarketDataSource
If pIndex = 0 Then Exit Function
Set getTickerFromIndex = mTickerTable(pIndex).DataSource
End Function

Private Function getTickerGridRow( _
                ByVal pTicker As IMarketDataSource) As Long
Const ProcName As String = "getTickerGridRow"
On Error GoTo Err

getTickerGridRow = getTickerGridRowFromIndex(getTickerIndex(pTicker))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTickerGridRowFromIndex( _
                ByVal pIndex As Long) As Long
Const ProcName As String = "getTickerGridRowFromIndex"
On Error GoTo Err

getTickerGridRowFromIndex = mTickerTable(pIndex).TickerGridRow

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTickerIndex( _
                ByVal pTicker As IMarketDataSource) As Long
Const ProcName As String = "getTickerIndex"
On Error GoTo Err

' allow for the fact that the first tickertable entry is not used - it is the
' terminator of the selected entries chain

getTickerIndex = pTicker.Handle + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTickerIndexFromGridRow(ByVal pRow As Long) As Long
Const ProcName As String = "getTickerIndexFromGridRow"
On Error GoTo Err

getTickerIndexFromGridRow = TickerGrid.RowData(pRow)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTickerNameColumnValue(ByVal pRow As Long) As String
getTickerNameColumnValue = TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.TickerName))
End Function

Private Sub insertBlankRow(ByVal pRow As Long)
Const ProcName As String = "insertBlankRow"
On Error GoTo Err

Dim lIndex As Long
Dim lTicker As IMarketDataSource

TickerGrid.InsertRow pRow

For Each lTicker In mTickers
    If getTickerGridRow(lTicker) >= pRow Then setTickerGridRow lTicker, getTickerGridRow(lTicker) + 1
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isAlphaNumeric(KeyAscii As Integer) As Boolean
If KeyAscii = 32 Then isAlphaNumeric = True: Exit Function
If KeyAscii < 48 Then isAlphaNumeric = False: Exit Function
If KeyAscii < 58 Then isAlphaNumeric = True: Exit Function
If KeyAscii < 65 Then isAlphaNumeric = False: Exit Function
If KeyAscii < 91 Then isAlphaNumeric = True: Exit Function
If KeyAscii < 97 Then isAlphaNumeric = False: Exit Function
If KeyAscii < 123 Then isAlphaNumeric = True: Exit Function
End Function

Private Function isRowOccupied( _
                ByVal pRow As Long) As Boolean
Const ProcName As String = "isRowOccupied"
On Error GoTo Err

isRowOccupied = (getTickerIndexFromGridRow(pRow) <> 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isRowOccupiedNonError( _
                ByVal pRow As Long) As Boolean
Const ProcName As String = "isRowOccupiedNonError"
On Error GoTo Err

If Not isRowOccupied(pRow) Then
Else
    Dim lTicker As IMarketDataSource
    Set lTicker = getTickerFromGridRow(pRow)
    isRowOccupiedNonError = lTicker.State <> MarketDataSourceStateError
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isTickerSelected( _
                ByVal pIndex As Long) As Boolean
Const ProcName As String = "isTickerSelected"
On Error GoTo Err

isTickerSelected = mSelectedTickers.Contains(getTickerFromIndex(pIndex))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadColumnMap()
Const ProcName As String = "loadColumnMap"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

If mConfig.GetSetting(ConfigSettingColumnMap) = "" Then
    setupColumnMap TickerGridColumns.MaxColumn
Else
    Dim ar() As String
    ar = Split(mConfig.GetSetting(ConfigSettingColumnMap), ",")
    
    Dim i As Long
    For i = 0 To UBound(ar)
        mColumnMap(i) = CLng(ar(i))
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub listenToTicker(ByVal pTicker As IMarketDataSource)
Const ProcName As String = "listenToTicker"
On Error GoTo Err

pTicker.AddStateChangeListener Me
pTicker.AddQuoteListener Me
pTicker.AddPriceChangeListener Me
pTicker.AddErrorListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub moveSelectedRows(ByVal pIncrement As Long)
Const ProcName As String = "moveSelectedRows"
On Error GoTo Err

Dim lTicker As IMarketDataSource

TickerGrid.Redraw = False

For Each lTicker In mSelectedTickers
    Dim lRow As Long
    lRow = getTickerGridRowFromIndex(getTickerIndex(lTicker))
    MoveRow lRow, lRow + pIncrement
    adjustMovedRow lRow, lRow + pIncrement
Next

TickerGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub notifyTickerSymbolEntry()
Const ProcName As String = "notifyTickerSymbolEntry"
On Error GoTo Err

Dim lSymbol As String
lSymbol = getTickerNameColumnValue(mTickerSymbolRow)
If lSymbol <> "" Then RaiseEvent TickerSymbolEntered(lSymbol, mTickerSymbolRow)
stopEnteringTickerSymbol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processAlphaNumericKey(KeyAscii As Integer)
Const ProcName As String = "processAlphaNumericKey"
On Error GoTo Err

If TickerGrid.Row = 0 Then Exit Sub

If Not mEnteringTickerSymbol Then
    If Not startEnteringTickerSymbol Then Exit Sub
End If

setTickerNameColumnValue mTickerSymbolRow, getTickerNameColumnValue(mTickerSymbolRow) & UCase$(Chr$(KeyAscii))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTickerState(ByVal pDataSource As IMarketDataSource)
Const ProcName As String = "processTickerState"
On Error GoTo Err

Dim lIndex As Long
lIndex = getTickerIndex(pDataSource)

Dim lRow As Long
lRow = getTickerGridRowFromIndex(lIndex)
    
Select Case pDataSource.State
Case MarketDataSourceStates.MarketDataSourceStateCreated
    'setTickerNameColumnValue lRow,  "Starting"
Case MarketDataSourceStates.MarketDataSourceStateReady
    setTickerFields lRow, gGetContractFromContractFuture(pDataSource.ContractFuture)
    setFieldsHaveBeenSet lIndex
    storeTickerSettings pDataSource
Case MarketDataSourceStates.MarketDataSourceStateError
    If Not getFieldsHaveBeenSet(lIndex) Then
        If pDataSource.ContractFuture.IsAvailable Then
            setTickerFields lRow, gGetContractFromContractFuture(pDataSource.ContractFuture)
            setFieldsHaveBeenSet lIndex
        End If
    End If
    setRowError lRow, pDataSource.ErrorMessage
Case MarketDataSourceStates.MarketDataSourceStateStopped, MarketDataSourceStates.MarketDataSourceStateFinished
    ' if the DataSource was stopped by the application via a call to IMarketDataSource.Finish (rather
    ' than via this control), the entry will still be in the grid so Remove it
    If Not getTickerFromIndex(lIndex) Is Nothing Then removeTickerFromGrid pDataSource, True
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RefreshPriceChange()
Const ProcName As String = "refreshPriceChange"
On Error GoTo Err

If Not mRefreshPriceChangeTC Is Nothing Then Exit Sub

Redraw = False

Dim lTask As New PriceChangeRefreshTask
lTask.Initialise cloneTickers, Me
Set mRefreshPriceChangeTC = StartTask(lTask, PriorityLow)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub RefreshQuotes()
Const ProcName As String = "refreshQuotes"
On Error GoTo Err

If Not mRefreshQuotesTC Is Nothing Then Exit Sub

Redraw = False

Dim lTask As New QuotesRefreshTask
lTask.Initialise cloneTickers, Me
Set mRefreshQuotesTC = StartTask(lTask, PriorityLow)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function removeTickerFromGrid( _
                ByVal pTicker As IMarketDataSource, _
                ByVal pRemoveGridRow As Boolean) As IMarketDataSource
Const ProcName As String = "removeTickerFromGrid"
On Error GoTo Err

Dim lRow As Long
lRow = getTickerGridRow(pTicker)

removeTickerFromConfig pTicker
mTickers.Remove pTicker.Key

Dim lIndex As Long
lIndex = getTickerIndex(pTicker)
Set mTickerTable(lIndex).DataSource = Nothing
setTickerGridRowFromIndex lIndex, 0

mSelectedTickers.Remove pTicker

If pTicker.State <> MarketDataSourceStateFinished Then stopListeningToTicker pTicker

If pRemoveGridRow Then
    TickerGrid.RemoveItem lRow

    Dim lTicker As IMarketDataSource
    For Each lTicker In mTickers
        If getTickerGridRow(lTicker) > lRow Then setTickerGridRow lTicker, getTickerGridRow(lTicker) - 1
    Next
Else
    clearGridRow lRow
End If

Set removeTickerFromGrid = pTicker

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeTickerFromConfig( _
                ByVal pTicker As IMarketDataSource)
Const ProcName As String = "removeTickerFromConfig"
On Error GoTo Err

If mTickersConfigSection Is Nothing Then Exit Sub
If pTicker.IsTickReplay Then Exit Sub

If mIteratingTickersConfig Then
    DeferAction Me, pTicker, 1000
    Exit Sub
End If

Dim tickerConfigSection As ConfigurationSection
Set tickerConfigSection = mTickersConfigSection.GetConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
If Not tickerConfigSection Is Nothing Then tickerConfigSection.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub selectATickerByIndex( _
                ByVal pIndex As Long)
Const ProcName As String = "selectATickerByIndex"
On Error GoTo Err

Dim lTicker As IMarketDataSource
Set lTicker = getTickerFromIndex(pIndex)
If lTicker Is Nothing Then Exit Sub
'If lTicker.State = MarketDataSourceStateError Then Exit Sub

mSelectedTickers.Add lTicker
toggleRowHighlight getTickerGridRowFromIndex(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub selectATickerByRow( _
                ByVal pRow As Long)
Const ProcName As String = "selectATickerByRow"
On Error GoTo Err

selectATickerByIndex getTickerIndexFromGridRow(pRow)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setColumnWidth( _
                ByVal pCol As Long, _
                ByVal widthChars As Long, _
                ByVal isLetters As Boolean)
Const ProcName As String = "setColumnWidth"
On Error GoTo Err

TickerGrid.ColWidth(pCol) = IIf(isLetters, mLetterWidth, mDigitWidth) * widthChars

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setColumnWidths()
Const ProcName As String = "setColumnWidths"
On Error GoTo Err

TickerGrid.Redraw = False
setColumnWidth TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, True
setColumnWidth TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, True
setColumnWidth TickerGridColumns.CurrencyCode, TickerGridColumnWidths.CurrencyWidth, True
setColumnWidth TickerGridColumns.BidSize, TickerGridColumnWidths.BidSizeWidth, False
setColumnWidth TickerGridColumns.Bid, TickerGridColumnWidths.BidWidth, False
setColumnWidth TickerGridColumns.Ask, TickerGridColumnWidths.AskWidth, False
setColumnWidth TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, False
setColumnWidth TickerGridColumns.Trade, TickerGridColumnWidths.TradeWidth, False
setColumnWidth TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSizeWidth, False
setColumnWidth TickerGridColumns.Volume, TickerGridColumnWidths.VolumeWidth, False
setColumnWidth TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, False
setColumnWidth TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, False
setColumnWidth TickerGridColumns.HighPrice, TickerGridColumnWidths.HighWidth, False
setColumnWidth TickerGridColumns.LowPrice, TickerGridColumnWidths.LowWidth, False
setColumnWidth TickerGridColumns.OpenPrice, TickerGridColumnWidths.OpenWidth, False
setColumnWidth TickerGridColumns.ClosePrice, TickerGridColumnWidths.CloseWidth, False
setColumnWidth TickerGridColumns.OpenInterest, TickerGridColumnWidths.OpenInterestWidth, False
setColumnWidth TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, True
setColumnWidth TickerGridColumns.Symbol, TickerGridColumnWidths.SymbolWidth, True
setColumnWidth TickerGridColumns.secType, TickerGridColumnWidths.SecTypeWidth, True
setColumnWidth TickerGridColumns.Expiry, TickerGridColumnWidths.ExpiryWidth, True
setColumnWidth TickerGridColumns.Exchange, TickerGridColumnWidths.ExchangeWidth, True
setColumnWidth TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, True
setColumnWidth TickerGridColumns.Strike, TickerGridColumnWidths.StrikeWidth, False
setColumnWidth TickerGridColumns.ErrorText, TickerGridColumnWidths.ErrorTextWidth, True
TickerGrid.Redraw = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setFieldsHaveBeenSet(ByVal pIndex As Long)
mTickerTable(pIndex).FieldsHaveBeenSet = True
End Sub

Private Sub clearGridRow(ByVal pRow As Long)
Const ProcName As String = "clearGridRow"
On Error GoTo Err

If pRow < 0 Then Exit Sub

Dim i As Long
For i = 1 To TickerGrid.Cols - 1
    TickerGrid.BeginCellEdit pRow, i
    TickerGrid.CellBackColor = 0
    TickerGrid.CellForeColor = 0
    TickerGrid.CellFontBold = False
    TickerGrid.EndCellEdit
    TickerGrid.TextMatrix(pRow, i) = ""
Next

TickerGrid.RowData(pRow) = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setRowError(ByVal pRow As Long, ByVal pErrorMessage As String)
Const ProcName As String = "setRowError"
On Error GoTo Err

If pRow < 0 Then Exit Sub

Dim i As Long
For i = 1 To TickerGrid.Cols - 1
    TickerGrid.BeginCellEdit pRow, i
    TickerGrid.CellBackColor = CErroredRowBackColor
    TickerGrid.CellForeColor = CErroredRowForeColor
    TickerGrid.CellFontBold = True
    TickerGrid.EndCellEdit
Next

TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.ErrorText)) = pErrorMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTickerFields(ByVal pRow As Long, ByVal pContract As IContract)
Const ProcName As String = "setTickerFields"
On Error GoTo Err

TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.CurrencyCode)) = pContract.Specifier.CurrencyCode
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.Description)) = pContract.Description
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.Exchange)) = pContract.Specifier.Exchange
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.Expiry)) = IIf(pContract.ExpiryDate = 0, "", pContract.ExpiryDate)
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.OptionRight)) = OptionRightToString(pContract.Specifier.Right)
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.secType)) = SecTypeToString(pContract.Specifier.secType)
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.Strike)) = pContract.Specifier.Strike
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.Symbol)) = pContract.Specifier.Symbol
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.TickerName)) = pContract.Specifier.LocalSymbol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTickerGridRow( _
                ByVal pTicker As IMarketDataSource, _
                ByVal pRow As Long)
Const ProcName As String = "setTickerGridRow"
On Error GoTo Err

setTickerGridRowFromIndex getTickerIndex(pTicker), pRow
storeTickerSettings pTicker

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTickerGridRowFromIndex( _
                ByVal pIndex As Long, _
                ByVal pRow As Long)
Const ProcName As String = "setTickerGridRowFromIndex"
On Error GoTo Err

If pIndex = 0 Then Exit Sub
mTickerTable(pIndex).TickerGridRow = pRow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setTickerIndexForRow(ByVal pRow As Long, ByVal pTickerIndex As Long)
TickerGrid.RowData(pRow) = pTickerIndex
End Sub

Private Sub setTickerNameColumnValue(ByVal pRow As Long, ByVal pValue As String)
TickerGrid.TextMatrix(pRow, mColumnMap(TickerGridColumns.TickerName)) = pValue
End Sub

Private Sub setupColumnMap( _
                    ByVal maxIndex As Long)
Const ProcName As String = "setupColumnMap"
On Error GoTo Err

ReDim mColumnMap(maxIndex) As Long

Dim i As Long
For i = 0 To UBound(mColumnMap)
    mColumnMap(i) = i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupDefaultTickerGridColumns()
Const ProcName As String = "setupDefaultTickerGridColumns"
On Error GoTo Err

gLogger.Log "Setting up default ticker grid columns", ProcName, ModuleName, LogLevelDetail

TickerGrid.Redraw = False

setupTickerGridColumn TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.CurrencyCode, TickerGridColumnWidths.CurrencyWidth, True, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.BidSize, TickerGridColumnWidths.BidSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Bid, TickerGridColumnWidths.BidWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Ask, TickerGridColumnWidths.AskWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Trade, TickerGridColumnWidths.TradeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Volume, TickerGridColumnWidths.VolumeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.HighPrice, TickerGridColumnWidths.HighWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.LowPrice, TickerGridColumnWidths.LowWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.OpenPrice, TickerGridColumnWidths.OpenWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.ClosePrice, TickerGridColumnWidths.CloseWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.OpenInterest, TickerGridColumnWidths.OpenInterestWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.Symbol, TickerGridColumnWidths.SymbolWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.secType, TickerGridColumnWidths.SecTypeWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.Expiry, TickerGridColumnWidths.ExpiryWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.Exchange, TickerGridColumnWidths.ExchangeWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.Strike, TickerGridColumnWidths.StrikeWidth, False, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.ErrorText, TickerGridColumnWidths.ErrorTextWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter

TickerGrid.Redraw = True

gLogger.Log "Default ticker grid columns setup completed", ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupDefaultTickerGridHeaders()
Const ProcName As String = "setupDefaultTickerGridHeaders"
On Error GoTo Err

setupTickerGridHeader TickerGridColumns.Selector, ""
setupTickerGridHeader TickerGridColumns.TickerName, "Name"
setupTickerGridHeader TickerGridColumns.CurrencyCode, "Curr"
setupTickerGridHeader TickerGridColumns.BidSize, "Bid Size"
setupTickerGridHeader TickerGridColumns.Bid, "Bid"
setupTickerGridHeader TickerGridColumns.Ask, "Ask"
setupTickerGridHeader TickerGridColumns.AskSize, "Ask Size"
setupTickerGridHeader TickerGridColumns.Trade, "Last"
setupTickerGridHeader TickerGridColumns.TradeSize, "Last Size"
setupTickerGridHeader TickerGridColumns.Volume, "Volume"
setupTickerGridHeader TickerGridColumns.Change, "Chg"
setupTickerGridHeader TickerGridColumns.ChangePercent, "Chg %"
setupTickerGridHeader TickerGridColumns.HighPrice, "High"
setupTickerGridHeader TickerGridColumns.LowPrice, "Low"
setupTickerGridHeader TickerGridColumns.OpenPrice, "Open"
setupTickerGridHeader TickerGridColumns.ClosePrice, "Close"
setupTickerGridHeader TickerGridColumns.OpenInterest, "Open interest"
setupTickerGridHeader TickerGridColumns.Description, "Description"
setupTickerGridHeader TickerGridColumns.Symbol, "Symbol"
setupTickerGridHeader TickerGridColumns.secType, "Sec Type"
setupTickerGridHeader TickerGridColumns.Expiry, "Expiry"
setupTickerGridHeader TickerGridColumns.Exchange, "Exchange"
setupTickerGridHeader TickerGridColumns.OptionRight, "Right"
setupTickerGridHeader TickerGridColumns.Strike, "Strike"
setupTickerGridHeader TickerGridColumns.ErrorText, "Error message"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSummaryTickerGridColumns()
Const ProcName As String = "setupSummaryTickerGridColumns"
On Error GoTo Err

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
    
setupColumnMap TickerGridSummaryColumns.MaxSummaryColumn

setupTickerGridColumn TickerGridSummaryColumns.Selector, TickerGridSummaryColumnWidths.SelectorWidth, True, TWControls40.AlignmentSettings.TwGridAlignCenterBottom
setupTickerGridColumn TickerGridSummaryColumns.TickerName, TickerGridSummaryColumnWidths.NameWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.BidSize, TickerGridSummaryColumnWidths.BidSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Bid, TickerGridSummaryColumnWidths.BidWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Ask, TickerGridSummaryColumnWidths.AskWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.AskSize, TickerGridSummaryColumnWidths.AskSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Trade, TickerGridSummaryColumnWidths.TradeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.TradeSize, TickerGridSummaryColumnWidths.TradeSizeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Volume, TickerGridSummaryColumnWidths.VolumeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Change, TickerGridSummaryColumnWidths.ChangeWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.ChangePercent, TickerGridSummaryColumnWidths.ChangePercentWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.OpenInterest, TickerGridSummaryColumnWidths.OpenInterestWidth, False, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.ErrorText, TickerGridSummaryColumnWidths.ErrorTextWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSummaryTickerGridHeaders()
Const ProcName As String = "setupSummaryTickerGridHeaders"
On Error GoTo Err

setupTickerGridHeader TickerGridSummaryColumns.Selector, ""
setupTickerGridHeader TickerGridSummaryColumns.TickerName, "Name"
setupTickerGridHeader TickerGridSummaryColumns.BidSize, "Bid Size"
setupTickerGridHeader TickerGridSummaryColumns.Bid, "Bid"
setupTickerGridHeader TickerGridSummaryColumns.Ask, "Ask"
setupTickerGridHeader TickerGridSummaryColumns.AskSize, "Ask Size"
setupTickerGridHeader TickerGridSummaryColumns.Trade, "Last"
setupTickerGridHeader TickerGridSummaryColumns.TradeSize, "Last Size"
setupTickerGridHeader TickerGridSummaryColumns.Volume, "Volume"
setupTickerGridHeader TickerGridSummaryColumns.Change, "Chg"
setupTickerGridHeader TickerGridSummaryColumns.ChangePercent, "Chg %"
setupTickerGridHeader TickerGridSummaryColumns.OpenInterest, "Open interest"
setupTickerGridHeader TickerGridSummaryColumns.ErrorText, "Error message"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupTickerGridColumn( _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Long, _
                ByVal isLetters As Boolean, _
                ByVal align As TWControls40.AlignmentSettings)
Const ProcName As String = "setupTickerGridColumn"
On Error GoTo Err

columnNumber = mColumnMap(columnNumber)

With TickerGrid
    
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .ColWidth(columnNumber) = 0
    End If
    
    .ColData(columnNumber) = columnNumber
    
    setColumnWidth columnNumber, columnWidth, isLetters
    
    .ColAlignment(columnNumber) = align
    .ColAlignmentFixed(columnNumber) = TwGridAlignCenterCenter
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupTickerGridHeader( _
                ByVal columnNumber As Long, _
                ByVal pHeading As String)
Const ProcName As String = "setupTickerGridHeader"
On Error GoTo Err

TickerGrid.TextMatrix(0, mColumnMap(columnNumber)) = pHeading

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function startEnteringTickerSymbol() As Boolean
Const ProcName As String = "startEnteringTickerSymbol"
On Error GoTo Err

If isRowOccupiedNonError(TickerGrid.Row) Then Exit Function

mTickerSymbolRow = TickerGrid.Row
mEnteringTickerSymbol = True

If isRowOccupied(mTickerSymbolRow) Then
    Dim lName As String
    lName = getTickerNameColumnValue(mTickerSymbolRow)
    
    RaiseEvent ErroredTickerRemoved(removeTickerFromGrid(getTickerFromGridRow(TickerGrid.Row), False))
    
    setTickerNameColumnValue mTickerSymbolRow, lName
End If

startEnteringTickerSymbol = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub stopEnteringTickerSymbol()
Const ProcName As String = "stopEnteringTickerSymbol"
On Error GoTo Err

If mEnteringTickerSymbol Then
    mEnteringTickerSymbol = False
    'setTickerNameColumnValue mTickerSymbolRow, ""
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub stopListeningToTicker(ByVal pTicker As IMarketDataSource)
Const ProcName As String = "stopListeningToTicker"
On Error GoTo Err

pTicker.RemoveStateChangeListener Me
pTicker.RemoveQuoteListener Me
pTicker.RemovePriceChangeListener Me
pTicker.RemoveErrorListener Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub stopTicker( _
                ByVal pTicker As IMarketDataSource)
Const ProcName As String = "stopTicker"
On Error GoTo Err

removeTickerFromGrid pTicker, True
If pTicker.IsMarketDataRequested Then pTicker.StopMarketData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeColumnMap()
Const ProcName As String = "storeColumnMap"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

Dim s As String

Dim i As Long
For i = 0 To UBound(mColumnMap)
    s = s & IIf(s = "", "", ", ") & CStr(mColumnMap(i))
Next

mConfig.SetSetting ConfigSettingColumnMap, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeTickerSettings( _
                ByVal pTicker As IMarketDataSource)
Const ProcName As String = "storeTickerSettings"
On Error GoTo Err

Dim tickerConfigSection As ConfigurationSection

If mTickersConfigSection Is Nothing Then Exit Sub
If pTicker.IsTickReplay Then Exit Sub
If Not pTicker.ContractFuture.IsAvailable Then Exit Sub
        
Set tickerConfigSection = mTickersConfigSection.AddConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
tickerConfigSection.SetSetting ConfigSettingRowIndex, getTickerGridRow(pTicker)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub toggleRowHighlight(ByVal pRow As Long)
Const ProcName As String = "toggleRowHighlight"
On Error GoTo Err

If pRow < 0 Then Exit Sub

Dim i As Long
For i = 1 To TickerGrid.Cols - 1
    TickerGrid.BeginCellEdit pRow, i
    If TickerGrid.CellFontBold Then
        TickerGrid.CellFontBold = False
    Else
        TickerGrid.CellFontBold = True
    End If
    TickerGrid.EndCellEdit
Next

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.TickerName)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.CurrencyCode)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.Description)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.Exchange)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.secType)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

TickerGrid.BeginCellEdit pRow, mColumnMap(TickerGridColumns.Symbol)
TickerGrid.InvertCellColors
TickerGrid.EndCellEdit

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub toggleRowSelection( _
                ByVal pRow As Long)
Const ProcName As String = "toggleRowSelection"
On Error GoTo Err

If Not isRowOccupied(pRow) Then Exit Sub

If isTickerSelected(getTickerIndexFromGridRow(pRow)) Then
    deselectATickerByRow pRow
Else
    selectATickerByRow pRow
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub truncateTickerSymbol()
Const ProcName As String = "truncateTickerSymbol"
On Error GoTo Err

If Not mEnteringTickerSymbol Then
    If Not startEnteringTickerSymbol Then Exit Sub
End If

Dim s As String
s = getTickerNameColumnValue(mTickerSymbolRow)
If s <> "" Then setTickerNameColumnValue mTickerSymbolRow, Left$(s, Len(s) - 1)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
