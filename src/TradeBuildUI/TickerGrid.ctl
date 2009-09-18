VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#39.0#0"; "TWControls10.ocx"
Begin VB.UserControl TickerGrid 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox FontPicture 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   555
      TabIndex        =   1
      Top             =   3120
      Visible         =   0   'False
      Width           =   615
   End
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
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event ColMoved(ByVal fromCol As Long, ByVal toCol As Long) 'MappingInfo=TickerGrid,TickerGrid,-1,ColMoved
Event ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean) 'MappingInfo=TickerGrid,TickerGrid,-1,ColMoving
Event DblClick() 'MappingInfo=TickerGrid,TickerGrid,-1,DblClick
Attribute DblClick.VB_UserMemId = -601
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyDown
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyPress
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=TickerGrid,TickerGrid,-1,KeyUp
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseDown
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseMove
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=TickerGrid,TickerGrid,-1,MouseUp
Attribute MouseUp.VB_UserMemId = -607
Event RowMoved(ByVal fromRow As Long, ByVal toRow As Long) 'MappingInfo=TickerGrid,TickerGrid,-1,RowMoved
Event RowMoving(ByVal fromRow As Long, ByVal toRow As Long, Cancel As Boolean) 'MappingInfo=TickerGrid,TickerGrid,-1,RowMoving
Event SelectionChanged()
Event TickerStarted(ByVal row As Long)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "TickerGrid"

Private Const ConfigSectionContractspecifier            As String = "ContractSpecifier"
Private Const ConfigSectionGrid                         As String = "Grid"
Private Const ConfigSectionTicker                       As String = "Ticker"
Private Const ConfigSectionTickers                      As String = "Tickers"

Private Const ConfigSettingHistorical                   As String = "&Historical"
Private Const ConfigSettingOptions                      As String = "&Options"

Private Const ConfigSettingPositiveChangeBackColor      As String = "&PositiveChangeBackColor"
Private Const ConfigSettingPositiveChangeForeColor      As String = "&PositiveChangeForeColor"
Private Const ConfigSettingNegativeChangeBackColor      As String = "&NegativeChangeBackColor"
Private Const ConfigSettingNegativeChangeForeColor      As String = "&NegativeChangeForeColor"
Private Const ConfigSettingIncreasedValueColor          As String = "&IncreasedValueColor"
Private Const ConfigSettingHighlightPriceChanges        As String = "&HighlightPriceChanges"
Private Const ConfigSettingDecreasedValueColor          As String = "&DecreasedValueColor"
Private Const ConfigSettingColumnMap                    As String = ".ColumnMap"

Private Const ConfigSettingContractSpecCurrency         As String = ConfigSectionContractspecifier & "&Currency"
Private Const ConfigSettingContractSpecExpiry           As String = ConfigSectionContractspecifier & "&Expiry"
Private Const ConfigSettingContractSpecExchange         As String = ConfigSectionContractspecifier & "&Exchange"
Private Const ConfigSettingContractSpecLocalSYmbol      As String = ConfigSectionContractspecifier & "&LocalSymbol"
Private Const ConfigSettingContractSpecRight            As String = ConfigSectionContractspecifier & "&Right"
Private Const ConfigSettingContractSpecSecType          As String = ConfigSectionContractspecifier & "&SecType"
Private Const ConfigSettingContractSpecStrikePrice      As String = ConfigSectionContractspecifier & "&StrikePrice"
Private Const ConfigSettingContractSpecSymbol           As String = ConfigSectionContractspecifier & "&Symbol"

Private Const GridRowsInitial As Long = 25
Private Const GridRowsIncrement As Long = 25

Private Const TickerTableEntriesInitial As Long = 4
Private Const TickerTableEntriesGrowthFactor As Long = 2

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
    currencyCode = 2
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
    symbol = 18
    secType = 19
    expiry = 20
    exchange = 21
    OptionRight = 22
    strike = 23
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
    BidSize
    Bid
    Ask
    AskSize
    Trade
    TradeSize
    Volume
    Change
    ChangePercent
    OpenInterest
    MaxSummaryColumn = OpenInterest
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
    theTicker               As Ticker
    tickerGridRow           As Long
'    nextSelected            As Long
'    prevSelected            As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mWorkspace As Workspace
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

Private WithEvents mSelectedTickers As SelectedTickers
Attribute mSelectedTickers.VB_VarHelpID = -1

Private mPositiveChangeBackColor As OLE_COLOR
Private mPositiveChangeForeColor As OLE_COLOR
Private mNegativeChangeBackColor As OLE_COLOR
Private mNegativeChangeForeColor As OLE_COLOR

Private mIncreasedValueColor As OLE_COLOR
Private mDecreasedValueColor As OLE_COLOR

Private mConfig As ConfigurationSection

Private mHighlightPriceChanges As Boolean

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()

ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
mNextGridRowIndex = 1

calcAverageCharacterWidths UserControl.Font

Set mSelectedTickers = New SelectedTickers

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
HighlightPriceChanges = True

TickerGrid.AllowBigSelection = True
AllowUserResizing = TwGridResizeBoth
RowSizingMode = TwGridRowSizeAll
FillStyle = TwGridFillRepeat
TickerGrid.FocusRect = TwGridFocusNone
TickerGrid.HighLight = TwGridHighlightNever
    
TickerGrid.Cols = 2
Rows = GridRowsInitial
TickerGrid.FixedRows = 1
TickerGrid.FixedCols = 1

setupDefaultTickerGridColumns

End Sub

Private Sub UserControl_Resize()
TickerGrid.Left = 0
TickerGrid.Top = 0
TickerGrid.Width = UserControl.Width
TickerGrid.Height = UserControl.Height
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

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
TickerGrid.FontFixed = PropBag.ReadProperty("FontFixed", UserControl.Ambient.Font)
TickerGrid.Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
TickerGrid.ForeColorSel = PropBag.ReadProperty("ForeColorSel", -2147483634)
TickerGrid.ForeColorFixed = PropBag.ReadProperty("ForeColorFixed", -2147483630)
TickerGrid.foreColor = PropBag.ReadProperty("ForeColor", &H80000008)
TickerGrid.FocusRect = PropBag.ReadProperty("FocusRect", TwGridFocusNone)
TickerGrid.FixedRows = PropBag.ReadProperty("FixedRows", 1)
TickerGrid.FixedCols = PropBag.ReadProperty("FixedCols", 1)
TickerGrid.FillStyle = PropBag.ReadProperty("FillStyle", TwGridFillRepeat)
TickerGrid.Cols = PropBag.ReadProperty("Cols", 2)

setupDefaultTickerGridColumns

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

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
Call PropBag.WriteProperty("FontFixed", TickerGrid.FontFixed)
Call PropBag.WriteProperty("Font", TickerGrid.Font)
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
Dim lTicker As Ticker
Dim failpoint As Long
On Error GoTo Err

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

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="PriceChangeListener_Change", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' QuoteListener Interface Members
'@================================================================================

Private Sub QuoteListener_ask(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Ask)
displaySize ev, mColumnMap(TickerGridColumns.AskSize)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_ask", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_bid(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Bid)
displaySize ev, mColumnMap(TickerGridColumns.BidSize)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_bid", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_high(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.HighPrice)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_high", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_Low(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.LowPrice)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_Low", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_openInterest(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displaySize ev, mColumnMap(TickerGridColumns.OpenInterest)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_openInterest", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_previousClose(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.ClosePrice)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_previousClose", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_sessionOpen(ev As TradeBuild26.QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.OpenPrice)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_sessionOpen", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_trade(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displayPrice ev, mColumnMap(TickerGridColumns.Trade)
displaySize ev, mColumnMap(TickerGridColumns.TradeSize)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_trade", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub QuoteListener_volume(ev As QuoteEvent)

Dim failpoint As Long
On Error GoTo Err

displaySize ev, mColumnMap(TickerGridColumns.Volume)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="QuoteListener_volume", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TickerGrid_ColMoved( _
                ByVal fromCol As Long, _
                ByVal toCol As Long)
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If fromCol < toCol Then
    For i = fromCol To toCol
        mColumnMap(TickerGrid.ColData(i)) = i
    Next
Else
    For i = toCol To fromCol
        mColumnMap(TickerGrid.ColData(i)) = i
    Next
End If

storeColumnMap

RaiseEvent ColMoved(fromCol, toCol)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_ColMoved", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_ColMoving(ByVal fromCol As Long, ByVal toCol As Long, Cancel As Boolean)
Dim failpoint As Long
On Error GoTo Err

    RaiseEvent ColMoving(fromCol, toCol, Cancel)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_ColMoving", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_Click()
Dim row As Long
Dim rowSel As Long
Dim col As Long
Dim colSel As Long
Dim failpoint As Long
On Error GoTo Err

row = TickerGrid.row
rowSel = TickerGrid.rowSel
col = TickerGrid.col
colSel = TickerGrid.colSel

mSelectedTickers.BeginChange
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
mSelectedTickers.EndChange

RaiseEvent Click

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_Click", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_DblClick()
Dim failpoint As Long
On Error GoTo Err

RaiseEvent DblClick

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_DblClick", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_KeyDown(KeyCode As Integer, Shift As Integer)
Dim failpoint As Long
On Error GoTo Err

RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_KeyDown", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_KeyPress(KeyAscii As Integer)
Dim failpoint As Long
On Error GoTo Err

RaiseEvent KeyPress(KeyAscii)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_KeyPress", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_KeyUp(KeyCode As Integer, Shift As Integer)
Dim failpoint As Long
On Error GoTo Err

RaiseEvent KeyUp(KeyCode, Shift)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_KeyUp", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Dim failpoint As Long
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_MouseDown", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

RaiseEvent MouseMove(Button, Shift, X, Y)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_MouseMove", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Dim failpoint As Long
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)
RaiseEvent MouseUp(Button, Shift, X, Y)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_MouseUp", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_RowMoved( _
                ByVal fromRow As Long, _
                ByVal toRow As Long)
Dim i As Long
Dim lTicker As Ticker

Dim failpoint As Long
On Error GoTo Err

Set lTicker = mTickerTable(TickerGrid.rowdata(toRow)).theTicker

If fromRow < toRow Then
    For i = fromRow To toRow
        mTickerTable(TickerGrid.rowdata(i)).tickerGridRow = i
        moveRowDownInConfig lTicker
    Next
Else
    For i = toRow To fromRow
        mTickerTable(TickerGrid.rowdata(i)).tickerGridRow = i
        moveRowUpInConfig lTicker
    Next
End If

RaiseEvent RowMoved(fromRow, toRow)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_RowMoved", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub TickerGrid_RowMoving( _
                ByVal fromRow As Long, _
                ByVal toRow As Long, _
                Cancel As Boolean)
Dim failpoint As Long
On Error GoTo Err

If toRow > mNextGridRowIndex Then
    Cancel = True
Else
    RaiseEvent RowMoving(fromRow, toRow, Cancel)
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="TickerGrid_RowMoving", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mCountTimer Event Handlers
'@================================================================================

Private Sub mCountTimer_TimerExpired()
mLogger.Log LogLevelMediumDetail, "TickerGrid: events per second=" & mEventCount / 10
mEventCount = 0
End Sub

'@================================================================================
' mSelectedTickers Event Handlers
'@================================================================================

Private Sub mSelectedTickers_SelectionChanged()
Dim failpoint As Long
On Error GoTo Err

RaiseEvent SelectionChanged

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="mSelectedTickers_SelectionChanged", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_StateChange(ev As StateChangeEvent)
Dim lTicker As Ticker
Dim index As Long
Dim lContract As Contract
    

Dim failpoint As Long
On Error GoTo Err

Set lTicker = ev.Source

index = getTickerIndex(lTicker)
    
Select Case ev.State
Case TickerStateCreated

Case TickerStateStarting
    
    If lTicker.IsHistorical Then Exit Sub
    
    Do While index > UBound(mTickerTable)
        ReDim Preserve mTickerTable((UBound(mTickerTable) + 1) * TickerTableEntriesGrowthFactor - 1) As TickerTableEntry
    Loop
    
    Set mTickerTable(index).theTicker = lTicker

    If mNextGridRowIndex > TickerGrid.Rows - 5 Then
        TickerGrid.Rows = TickerGrid.Rows + GridRowsIncrement
    End If
    
    mTickerTable(index).tickerGridRow = mNextGridRowIndex
    mNextGridRowIndex = mNextGridRowIndex + 1
    lTicker.AddQuoteListener Me
    lTicker.AddPriceChangeListener Me

    TickerGrid.row = mTickerTable(index).tickerGridRow
    TickerGrid.rowdata(mTickerTable(index).tickerGridRow) = index
    
    TickerGrid.col = mColumnMap(TickerGridColumns.TickerName)
    TickerGrid.Text = "Starting"
    
    gLogger.Log LogLevels.LogLevelNormal, ProjectName & "." & ModuleName & ": Added Ticker to grid " & lTicker.Key
    
Case TickerStateReady
    
    If lTicker.IsHistorical Then
        ' Add it to the config but don't map it into the grid
        addTickerToConfig lTicker
        Exit Sub
    End If
    
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
    
    TickerGrid.col = mColumnMap(TickerGridColumns.secType)
    TickerGrid.Text = SecTypeToString(lContract.specifier.secType)
    
    TickerGrid.col = mColumnMap(TickerGridColumns.strike)
    TickerGrid.Text = lContract.specifier.strike
    
    TickerGrid.col = mColumnMap(TickerGridColumns.symbol)
    TickerGrid.Text = lContract.specifier.symbol
    
    TickerGrid.col = mColumnMap(TickerGridColumns.TickerName)
    TickerGrid.Text = lContract.specifier.localSymbol
    
    addTickerToConfig lTicker
    
    RaiseEvent TickerStarted(mTickerTable(index).tickerGridRow)
    
Case TickerStateRunning
    
Case TickerStatePaused

Case TickerStateClosing

Case TickerStateStopped
    If lTicker.IsHistorical Then
        removeTickerFromConfig lTicker
        Exit Sub
    End If
    
    ' if the Ticker was stopped by the application via a call to Ticker.stopTicker (rather
    ' than via this control), the entry will still be in the grid so Remove it
    If Not mTickerTable(index).theTicker Is Nothing Then
        mTickerTable(index).theTicker.RemoveQuoteListener Me
        mTickerTable(index).theTicker.RemovePriceChangeListener Me
        removeTicker index
    End If
End Select

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="mTickers_StateChange", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Workspace() As Workspace
Attribute Workspace.VB_MemberFlags = "400"
Set Workspace = mWorkspace
End Property

Public Property Get SelectedTickers() As SelectedTickers
Attribute SelectedTickers.VB_MemberFlags = "400"
Set SelectedTickers = mSelectedTickers
End Property

Public Property Get ScrollBars() As TWControls10.ScrollBarsSettings
Attribute ScrollBars.VB_Description = "Specifies whether scroll bars are to be provided."
Attribute ScrollBars.VB_ProcData.VB_Invoke_Property = ";Behavior"
    ScrollBars = TickerGrid.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As TWControls10.ScrollBarsSettings)
    TickerGrid.ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

Public Property Get RowSizingMode() As TWControls10.RowSizingSettings
Attribute RowSizingMode.VB_Description = "Specifies whether resizing a row affects only that row or all rows."
Attribute RowSizingMode.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RowSizingMode = TickerGrid.RowSizingMode
End Property

Public Property Let RowSizingMode(ByVal New_RowSizingMode As TWControls10.RowSizingSettings)
    TickerGrid.RowSizingMode = New_RowSizingMode
    PropertyChanged "RowSizingMode"
End Property

Public Property Get Rows() As Long
Attribute Rows.VB_Description = "Specifies the initial number of rows (bear in mind that the header consumes one row)."
Attribute Rows.VB_ProcData.VB_Invoke_Property = ";Appearance"
    Rows = TickerGrid.Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
    TickerGrid.Rows = New_Rows
    PropertyChanged "Rows"
End Property

Public Property Get RowHeightMin() As Long
Attribute RowHeightMin.VB_Description = "Specifies the minimum height to which a row can be resized by the user."
Attribute RowHeightMin.VB_ProcData.VB_Invoke_Property = ";Behavior"
    RowHeightMin = TickerGrid.RowHeightMin
End Property

Public Property Let RowHeightMin(ByVal New_RowHeightMin As Long)
    TickerGrid.RowHeightMin = New_RowHeightMin
    PropertyChanged "RowHeightMin"
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
Attribute RowBackColorOdd.VB_Description = "Specifies the background color for odd-numbered rows."
Attribute RowBackColorOdd.VB_ProcData.VB_Invoke_Property = ";Appearance"
    RowBackColorOdd = TickerGrid.RowBackColorOdd
End Property

Public Property Let RowBackColorOdd(ByVal New_RowBackColorOdd As OLE_COLOR)
    TickerGrid.RowBackColorOdd = New_RowBackColorOdd
    PropertyChanged "RowBackColorOdd"
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
Attribute RowBackColorEven.VB_Description = "Specifies the background color for even-numbered rows."
Attribute RowBackColorEven.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute RowBackColorEven.VB_MemberFlags = "200"
    RowBackColorEven = TickerGrid.RowBackColorEven
End Property

Public Property Let RowBackColorEven(ByVal New_RowBackColorEven As OLE_COLOR)
    TickerGrid.RowBackColorEven = New_RowBackColorEven
    PropertyChanged "RowBackColorEven"
End Property

Public Property Get Redraw() As Boolean
Attribute Redraw.VB_MemberFlags = "400"
Redraw = TickerGrid.Redraw
End Property

Public Property Let Redraw(ByVal value As Boolean)
TickerGrid.Redraw = value
End Property

Public Property Let PositiveChangeBackColor(ByVal value As OLE_COLOR)
mPositiveChangeBackColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPositiveChangeBackColor, mPositiveChangeBackColor
If Not mTickers Is Nothing Then mTickers.RefreshPriceChange Me
PropertyChanged "PositiveChangeBackColor"
End Property

Public Property Get PositiveChangeBackColor() As OLE_COLOR
Attribute PositiveChangeBackColor.VB_Description = "Specifies the background color for price change cells when the price has increased."
Attribute PositiveChangeBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PositiveChangeBackColor = mPositiveChangeBackColor
End Property

Public Property Let PositiveChangeForeColor(ByVal value As OLE_COLOR)
mPositiveChangeForeColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPositiveChangeForeColor, mPositiveChangeForeColor
If Not mTickers Is Nothing Then mTickers.RefreshPriceChange Me
PropertyChanged "PositiveChangeForeColor"
End Property

Public Property Get PositiveChangeForeColor() As OLE_COLOR
Attribute PositiveChangeForeColor.VB_Description = "Specifies the foreground color for price change cells when the price has increased."
Attribute PositiveChangeForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PositiveChangeForeColor = mPositiveChangeForeColor
End Property

Public Property Let NegativeChangeBackColor(ByVal value As OLE_COLOR)
mNegativeChangeBackColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingNegativeChangeBackColor, mNegativeChangeBackColor
If Not mTickers Is Nothing Then mTickers.RefreshPriceChange Me
PropertyChanged "NegativeChangeBackColor"
End Property

Public Property Get NegativeChangeBackColor() As OLE_COLOR
Attribute NegativeChangeBackColor.VB_Description = "Specifies the background color for price change cells when the price has decreased."
Attribute NegativeChangeBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
NegativeChangeBackColor = mNegativeChangeBackColor
End Property

Public Property Let NegativeChangeForeColor(ByVal value As OLE_COLOR)
mNegativeChangeForeColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingNegativeChangeForeColor, mNegativeChangeForeColor
If Not mTickers Is Nothing Then mTickers.RefreshPriceChange Me
PropertyChanged "NegativeChangeForeColor"
End Property

Public Property Get NegativeChangeForeColor() As OLE_COLOR
Attribute NegativeChangeForeColor.VB_Description = "Specifies the foreground color for price change cells when the price has decreased."
Attribute NegativeChangeForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
NegativeChangeForeColor = mNegativeChangeForeColor
End Property

Public Property Let IncreasedValueColor(ByVal value As OLE_COLOR)
mIncreasedValueColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingIncreasedValueColor, mIncreasedValueColor
If Not mTickers Is Nothing Then mTickers.RefreshQuotes Me
PropertyChanged "IncreasedValueColor"
End Property

Public Property Get IncreasedValueColor() As OLE_COLOR
Attribute IncreasedValueColor.VB_Description = "Specifies the foreground color for price cells that have increased in value."
Attribute IncreasedValueColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
IncreasedValueColor = mIncreasedValueColor
End Property

Public Property Let HighlightPriceChanges(ByVal value As Boolean)
mHighlightPriceChanges = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHighlightPriceChanges, mHighlightPriceChanges
If Not mTickers Is Nothing Then mTickers.RefreshPriceChange Me
PropertyChanged "HighlightPriceChanges"
End Property

Public Property Get HighlightPriceChanges() As Boolean
Attribute HighlightPriceChanges.VB_ProcData.VB_Invoke_Property = ";Behavior"
HighlightPriceChanges = mHighlightPriceChanges
End Property

Public Property Get GridLineWidth() As Long
Attribute GridLineWidth.VB_Description = "Specifies the thickness of the grid lines."
Attribute GridLineWidth.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridLineWidth = TickerGrid.GridLineWidth
End Property

Public Property Let GridLineWidth(ByVal New_GridLineWidth As Long)
    TickerGrid.GridLineWidth = New_GridLineWidth
    PropertyChanged "GridLineWidth"
End Property

Public Property Get GridColorFixed() As OLE_COLOR
Attribute GridColorFixed.VB_Description = "Specifies the color of the header grid lines."
Attribute GridColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColorFixed = TickerGrid.GridColorFixed
End Property

Public Property Let GridColorFixed(ByVal New_GridColorFixed As OLE_COLOR)
    TickerGrid.GridColorFixed = New_GridColorFixed
    PropertyChanged "GridColorFixed"
End Property

Public Property Get GridColor() As OLE_COLOR
Attribute GridColor.VB_Description = "Specifies the color of the grid lines."
Attribute GridColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    GridColor = TickerGrid.GridColor
End Property

Public Property Let GridColor(ByVal New_GridColor As OLE_COLOR)
    TickerGrid.GridColor = New_GridColor
    PropertyChanged "GridColor"
End Property

Public Property Get ForeColorFixed() As OLE_COLOR
Attribute ForeColorFixed.VB_Description = "Specifies the foreground color for header cells."
Attribute ForeColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    ForeColorFixed = TickerGrid.ForeColorFixed
End Property

Public Property Let ForeColorFixed(ByVal New_ForeColorFixed As OLE_COLOR)
    TickerGrid.ForeColorFixed = New_ForeColorFixed
    PropertyChanged "ForeColorFixed"
End Property

Public Property Get foreColor() As OLE_COLOR
Attribute foreColor.VB_Description = "Specifies the foreground color for non-header cells."
Attribute foreColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute foreColor.VB_UserMemId = -513
    foreColor = TickerGrid.foreColor
End Property

Public Property Let foreColor(ByVal New_ForeColor As OLE_COLOR)
    TickerGrid.foreColor = New_ForeColor
    PropertyChanged "ForeColor"
End Property

Public Property Get FontFixed() As StdFont
Attribute FontFixed.VB_Description = "Specifies the font to be used for header cells."
Attribute FontFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
Set FontFixed = TickerGrid.FontFixed
End Property

Public Property Let FontFixed(ByVal value As StdFont)
Set TickerGrid.FontFixed = value
End Property

Public Property Set FontFixed(ByVal value As StdFont)
TickerGrid.FontFixed = value
End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Specifies the font to be used for non-header cells."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Font.VB_UserMemId = -512
Set Font = TickerGrid.Font
End Property

Public Property Set Font(ByVal value As StdFont)
Set TickerGrid.Font = value
calcAverageCharacterWidths value
setColumnWidths
End Property

Public Property Let Font(ByVal value As StdFont)
Set TickerGrid.Font = value
calcAverageCharacterWidths value
setColumnWidths
End Property

Public Property Let DecreasedValueColor(ByVal value As OLE_COLOR)
mDecreasedValueColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDecreasedValueColor, mDecreasedValueColor
If Not mTickers Is Nothing Then mTickers.RefreshQuotes Me
PropertyChanged "DecreasedValueColor"
End Property

Public Property Get DecreasedValueColor() As OLE_COLOR
Attribute DecreasedValueColor.VB_Description = "Specifies the foreground color for price cells that have decreased in value."
Attribute DecreasedValueColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
DecreasedValueColor = mDecreasedValueColor
End Property

Public Property Let ConfigurationSection( _
                ByVal config As ConfigurationSection)
Attribute ConfigurationSection.VB_MemberFlags = "400"
Set mConfig = config

mConfig.SetSetting ConfigSettingPositiveChangeBackColor, mPositiveChangeBackColor
mConfig.SetSetting ConfigSettingPositiveChangeForeColor, mPositiveChangeForeColor
mConfig.SetSetting ConfigSettingNegativeChangeBackColor, mNegativeChangeBackColor
mConfig.SetSetting ConfigSettingNegativeChangeForeColor, mNegativeChangeForeColor
mConfig.SetSetting ConfigSettingIncreasedValueColor, mIncreasedValueColor
mConfig.SetSetting ConfigSettingDecreasedValueColor, mDecreasedValueColor

storeColumnMap

TickerGrid.ConfigurationSection = mConfig.AddPrivateConfigurationSection(ConfigSectionGrid)
End Property

Public Property Get BackColorFixed() As OLE_COLOR
Attribute BackColorFixed.VB_Description = "Specifies the background color of the fixed cells (ie row and column headers)."
Attribute BackColorFixed.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorFixed = TickerGrid.BackColorFixed
End Property

Public Property Let BackColorFixed(ByVal New_BackColorFixed As OLE_COLOR)
    TickerGrid.BackColorFixed = New_BackColorFixed
    PropertyChanged "BackColorFixed"
End Property

Public Property Get BackColorBkg() As OLE_COLOR
Attribute BackColorBkg.VB_Description = "Specifies the color of the area behind the rows and columns."
Attribute BackColorBkg.VB_ProcData.VB_Invoke_Property = ";Appearance"
    BackColorBkg = TickerGrid.BackColorBkg
End Property

Public Property Let BackColorBkg(ByVal New_BackColorBkg As OLE_COLOR)
    TickerGrid.BackColorBkg = New_BackColorBkg
    PropertyChanged "BackColorBkg"
End Property

Public Property Get backColor() As OLE_COLOR
Attribute backColor.VB_UserMemId = -501
Attribute backColor.VB_MemberFlags = "400"
    backColor = TickerGrid.backColor
End Property

Public Property Let backColor(ByVal New_BackColor As OLE_COLOR)
    TickerGrid.backColor = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get AllowUserResizing() As TWControls10.AllowUserResizeSettings
Attribute AllowUserResizing.VB_Description = "Specifies whethe the user is allowed to change the size of columns and/or rows."
Attribute AllowUserResizing.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowUserResizing = TickerGrid.AllowUserResizing
End Property

Public Property Let AllowUserResizing(ByVal New_AllowUserResizing As TWControls10.AllowUserResizeSettings)
    TickerGrid.AllowUserResizing = New_AllowUserResizing
    PropertyChanged "AllowUserResizing"
End Property

Public Property Get AllowUserReordering() As TWControls10.AllowUserReorderSettings
Attribute AllowUserReordering.VB_Description = "Specifies whether the user is allowed to change the order of columns and/or rows."
Attribute AllowUserReordering.VB_ProcData.VB_Invoke_Property = ";Behavior"
    AllowUserReordering = TickerGrid.AllowUserReordering
End Property

Public Property Let AllowUserReordering(ByVal New_AllowUserReordering As TWControls10.AllowUserReorderSettings)
    TickerGrid.AllowUserReordering = New_AllowUserReordering
    PropertyChanged "AllowUserReordering"
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub deselectSelectedTickers()
Dim index As Long
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

For i = 1 To mNextGridRowIndex - 1
    index = TickerGrid.rowdata(i)
    If isTickerSelected(index) Then
        highlightRow mTickerTable(index).tickerGridRow
    End If
Next

mSelectedTickers.Clear

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="deselectSelectedTickers", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub deselectTicker( _
                ByVal index As Long)
Dim failpoint As Long
On Error GoTo Err

deselectATicker index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="deselectTicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'Public Sub ExtendSelection(ByVal row As Long, ByVal col As Long)
'    TickerGrid.ExtendSelection row, col
'End Sub

Public Sub Finish()
Dim lTicker As Ticker

On Error GoTo Err

For Each lTicker In mTickers
    lTicker.RemoveQuoteListener Me
    lTicker.RemovePriceChangeListener Me
Next

Set mTickers = Nothing
TickerGrid.Clear
ReDim mTickerTable(TickerTableEntriesInitial - 1) As TickerTableEntry
mNextGridRowIndex = 1
mSelectedTickers.Clear
If Not mCountTimer Is Nothing Then mCountTimer.StopTimer
Exit Sub
Err:
'ignore any errors
End Sub

Public Function GetColFromX(ByVal X As Long) As Long
Dim failpoint As Long
On Error GoTo Err

    GetColFromX = TickerGrid.GetColFromX(X)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="GetColFromX", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function GetRowFromY(ByVal Y As Long) As Long
Dim failpoint As Long
On Error GoTo Err

    GetRowFromY = TickerGrid.GetRowFromY(Y)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="GetRowFromY", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub InvertCellColors()
Dim failpoint As Long
On Error GoTo Err

TickerGrid.InvertCellColors

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="InvertCellColors", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Dim tickersConfigSection As ConfigurationSection
Dim tickerConfigSection As ConfigurationSection
Dim contractSpec As contractSpecifier

Dim localSymbol As String
Dim symbol As String
Dim exchange As String
Dim secType As SecurityTypes
Dim currencyCode As String
Dim expiry As String
Dim strikePrice As Double
Dim optRight As OptionRights

Dim lTicker As Ticker

Dim failpoint As Long
On Error GoTo Err

Set mConfig = config

TickerGrid.LoadFromConfig mConfig.AddPrivateConfigurationSection(ConfigSectionGrid)

loadColumnMap
setupDefaultTickerGridHeaders

If mConfig.GetSetting(ConfigSettingPositiveChangeBackColor) <> "" Then mPositiveChangeBackColor = mConfig.GetSetting(ConfigSettingPositiveChangeBackColor)
If mConfig.GetSetting(ConfigSettingPositiveChangeForeColor) <> "" Then mPositiveChangeForeColor = mConfig.GetSetting(ConfigSettingPositiveChangeForeColor)
If mConfig.GetSetting(ConfigSettingNegativeChangeBackColor) <> "" Then mNegativeChangeBackColor = mConfig.GetSetting(ConfigSettingNegativeChangeBackColor)
If mConfig.GetSetting(ConfigSettingNegativeChangeForeColor) <> "" Then mNegativeChangeForeColor = mConfig.GetSetting(ConfigSettingNegativeChangeForeColor)
If mConfig.GetSetting(ConfigSettingIncreasedValueColor) <> "" Then mIncreasedValueColor = mConfig.GetSetting(ConfigSettingIncreasedValueColor)
If mConfig.GetSetting(ConfigSettingHighlightPriceChanges) <> "" Then mHighlightPriceChanges = mConfig.GetSetting(ConfigSettingHighlightPriceChanges)
If mConfig.GetSetting(ConfigSettingDecreasedValueColor) <> "" Then mDecreasedValueColor = mConfig.GetSetting(ConfigSettingDecreasedValueColor)

Set tickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)

For Each tickerConfigSection In tickersConfigSection
    With tickerConfigSection
        localSymbol = .GetSetting(ConfigSettingContractSpecLocalSYmbol, "")
        symbol = .GetSetting(ConfigSettingContractSpecSymbol, "")
        exchange = .GetSetting(ConfigSettingContractSpecExchange, "")
        secType = SecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, ""))
        currencyCode = .GetSetting(ConfigSettingContractSpecCurrency, "")
        expiry = .GetSetting(ConfigSettingContractSpecExpiry, "")
        strikePrice = CDbl("0" & .GetSetting(ConfigSettingContractSpecStrikePrice, "0.0"))
        optRight = OptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, ""))
        
        Set contractSpec = CreateContractSpecifier(localSymbol, _
                                                symbol, _
                                                exchange, _
                                                secType, _
                                                currencyCode, _
                                                expiry, _
                                                strikePrice, _
                                                optRight)
    End With
    
    Set lTicker = mTickers.Add(tickerConfigSection.GetSetting(ConfigSettingOptions), _
                                tickerConfigSection.InstanceQualifier)
    If tickerConfigSection.GetSetting(ConfigSettingHistorical, "False") Then
        lTicker.LoadTicker contractSpec
        gLogger.Log LogLevels.LogLevelNormal, ProjectName & "." & ModuleName & ": Loaded Ticker " & contractSpec.toString
    Else
        lTicker.StartTicker contractSpec
        gLogger.Log LogLevels.LogLevelNormal, ProjectName & "." & ModuleName & ": Started Ticker " & contractSpec.toString
    End If
Next

gLogger.Log LogLevels.LogLevelNormal, ProjectName & "." & ModuleName & ": Ticker grid loaded from config"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="LoadFromConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub monitorWorkspace( _
                ByVal pWorkspace As Workspace)
Dim failpoint As Long
On Error GoTo Err

If Not mTickers Is Nothing Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                            "TradeBuildUI.TickerGrid::monitorWorkspace", _
                                            "A workspace is already being monitored"
Set mWorkspace = pWorkspace
Set mTickers = pWorkspace.Tickers

Set mLogger = GetLogger("diag.tradebuild.tradebuildui")

Set mCountTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds, 10000)
mCountTimer.StartTimer

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="monitorWorkspace", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveFromConfig()
Dim failpoint As Long
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="RemoveFromConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub selectAllTickers()
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

mSelectedTickers.BeginChange
For i = 1 To mNextGridRowIndex - 1
    selectATicker i
Next
mSelectedTickers.EndChange

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="selectAllTickers", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub selectTicker( _
                ByVal row As Long)
Dim failpoint As Long
On Error GoTo Err

selectATicker row

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="selectTicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub setCellAlignment(ByVal row As Long, ByVal col As Long, pAlign As TWControls10.AlignmentSettings)
Dim failpoint As Long
On Error GoTo Err

TickerGrid.setCellAlignment row, col, pAlign

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setCellAlignment", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub stopAllTickers()
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

TickerGrid.Redraw = False

' do this in reverse order - most efficient when all tickers are being stopped
For i = TickerGrid.Rows - 1 To 1 Step -1
    If TickerGrid.rowdata(i) <> 0 Then
        stopTicker TickerGrid.rowdata(i)
    End If
Next
TickerGrid.Redraw = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="stopAllTickers", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub stopSelectedTickers()
Dim lTicker As Ticker

Dim failpoint As Long
On Error GoTo Err

TickerGrid.Redraw = False

For Each lTicker In mSelectedTickers
    stopTicker getTickerIndex(lTicker)
Next

TickerGrid.Redraw = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="stopSelectedTickers", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addTickerToConfig( _
                ByVal pTicker As Ticker)
Dim tickersConfigSection As ConfigurationSection
Dim tickerConfigSection As ConfigurationSection
Dim contractSpec As contractSpecifier

Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then
    If Not pTicker.ReplayingTickfile And _
        Not pTicker.Contract Is Nothing _
    Then
        Set tickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)
        Set tickerConfigSection = tickersConfigSection.AddConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
        tickerConfigSection.AddConfigurationSection ConfigSectionContractspecifier
        Set contractSpec = pTicker.Contract.specifier
        With tickerConfigSection
            .SetSetting ConfigSettingOptions, pTicker.Options
            .SetSetting ConfigSettingHistorical, CStr(pTicker.IsHistorical)
            
            .SetSetting ConfigSettingContractSpecLocalSYmbol, contractSpec.localSymbol
            .SetSetting ConfigSettingContractSpecSymbol, contractSpec.symbol
            .SetSetting ConfigSettingContractSpecExchange, contractSpec.exchange
            .SetSetting ConfigSettingContractSpecSecType, SecTypeToString(contractSpec.secType)
            .SetSetting ConfigSettingContractSpecCurrency, contractSpec.currencyCode
            .SetSetting ConfigSettingContractSpecExpiry, contractSpec.expiry
            .SetSetting ConfigSettingContractSpecStrikePrice, contractSpec.strike
            .SetSetting ConfigSettingContractSpecRight, OptionRightToString(contractSpec.Right)
        End With
        gLogger.Log LogLevels.LogLevelNormal, ProjectName & "." & ModuleName & ": Added Ticker to config " & pTicker.Key & "={" & pTicker.Contract.specifier.toString & "}"
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="addTickerToConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub calcAverageCharacterWidths( _
                ByVal afont As StdFont)
Dim failpoint As Long
On Error GoTo Err

mLetterWidth = getAverageCharacterWidth("ABCDEFGH IJKLMNOP QRST UVWX YZ", afont)
mDigitWidth = getAverageCharacterWidth(".0123456789", afont)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="calcAverageCharacterWidths", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub deselectATicker( _
                ByVal row As Long)
Dim index As Long
Dim failpoint As Long
On Error GoTo Err

index = TickerGrid.rowdata(row)
If isTickerSelected(index) Then
    mSelectedTickers.Remove mTickerTable(index).theTicker
    highlightRow mTickerTable(index).tickerGridRow
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="deselectATicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub displayPrice( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As Ticker
Dim failpoint As Long
On Error GoTo Err

Set lTicker = ev.Source
TickerGrid.row = getTickerGridRow(lTicker)
TickerGrid.col = col
TickerGrid.Text = GetFormattedPriceFromQuoteEvent(ev)

If ev.PriceChange = ValueChangeNone Or (Not mHighlightPriceChanges) Then
    If ev.recentPriceChange = ValueChangeUp Then
        TickerGrid.CellForeColor = mIncreasedValueColor
    ElseIf ev.recentPriceChange = ValueChangeDown Then
        TickerGrid.CellForeColor = mDecreasedValueColor
    Else
        TickerGrid.CellForeColor = 0
    End If
    
    TickerGrid.CellBackColor = 0
Else
    TickerGrid.CellBackColor = 0    ' reset backcolor to default
    TickerGrid.CellForeColor = TickerGrid.CellBackColor
    If ev.PriceChange = ValueChangeUp Then
        TickerGrid.CellBackColor = mIncreasedValueColor
    ElseIf ev.PriceChange = ValueChangeDown Then
        TickerGrid.CellBackColor = mDecreasedValueColor
    End If
End If

incrementEventCount

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="displayPrice", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub displaySize( _
                ev As QuoteEvent, _
                ByVal col As Long)
Dim lTicker As Ticker
Dim failpoint As Long
On Error GoTo Err

Set lTicker = ev.Source
TickerGrid.row = getTickerGridRow(lTicker)
TickerGrid.col = col
TickerGrid.Text = ev.size

If ev.sizeChange = ValueChangeNone Then
    If ev.recentSizeChange = ValueChangeUp Then
        TickerGrid.CellForeColor = mIncreasedValueColor
    ElseIf ev.recentSizeChange = ValueChangeDown Then
        TickerGrid.CellForeColor = mDecreasedValueColor
    Else
        TickerGrid.CellForeColor = 0
    End If
Else
    If ev.sizeChange = ValueChangeUp Then
        TickerGrid.CellForeColor = mIncreasedValueColor
    ElseIf ev.sizeChange = ValueChangeDown Then
        TickerGrid.CellForeColor = mDecreasedValueColor
    End If
End If

incrementEventCount

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="displaySize", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Function getAverageCharacterWidth( _
                ByVal widthString As String, _
                ByVal pFont As StdFont) As Long
Dim failpoint As Long
On Error GoTo Err

Set FontPicture.Font = pFont
getAverageCharacterWidth = FontPicture.TextWidth(widthString) / Len(widthString)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="getAverageCharacterWidth", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function getTickerGridRow( _
                ByVal pTicker As Ticker) As Long
Dim failpoint As Long
On Error GoTo Err

getTickerGridRow = mTickerTable(getTickerIndex(pTicker)).tickerGridRow

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="getTickerGridRow", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function getTickerIndex( _
                ByVal pTicker As Ticker) As Long
' allow for the fact that the first tickertable entry is not used - it is the
' terminator of the selected entries chain
Dim failpoint As Long
On Error GoTo Err

getTickerIndex = pTicker.Handle + 1

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="getTickerIndex", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub incrementEventCount()
mEventCount = mEventCount + 1
End Sub

Private Sub highlightRow(ByVal rowNumber As Long)
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

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

TickerGrid.col = mColumnMap(TickerGridColumns.secType)
TickerGrid.InvertCellColors

TickerGrid.col = mColumnMap(TickerGridColumns.symbol)
TickerGrid.InvertCellColors

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="highlightRow", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Function isRowSelected( _
                ByVal row As Long)
Dim failpoint As Long
On Error GoTo Err

isRowSelected = isTickerSelected(TickerGrid.rowdata(row))

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="isRowSelected", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function isTickerSelected( _
                ByVal index As Long)
Dim failpoint As Long
On Error GoTo Err

isTickerSelected = mSelectedTickers.Contains(mTickerTable(index).theTicker)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="isTickerSelected", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Sub loadColumnMap()
Dim ar() As String
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

If mConfig.GetSetting(ConfigSettingColumnMap) = "" Then
    setupColumnMap TickerGridColumns.MaxColumn
Else
    
    ar = Split(mConfig.GetSetting(ConfigSettingColumnMap), ",")
    
    ReDim mColumnMap(UBound(ar)) As Long
    
    For i = 0 To UBound(ar)
        mColumnMap(i) = CLng(ar(i))
    Next
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="loadColumnMap", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub moveRowDownInConfig( _
                ByVal pTicker As Ticker)
Dim tickersConfigSection As ConfigurationSection
Dim tickerConfigSection As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then
    If Not pTicker.ReplayingTickfile And _
        Not pTicker.Contract Is Nothing _
    Then
        Set tickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)
        Set tickerConfigSection = tickersConfigSection.GetConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
        If Not tickerConfigSection Is Nothing Then tickerConfigSection.MoveDown
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="moveRowDownInConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Private Sub moveRowUpInConfig( _
                ByVal pTicker As Ticker)
Dim tickersConfigSection As ConfigurationSection
Dim tickerConfigSection As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then
    If Not pTicker.ReplayingTickfile And _
        Not pTicker.Contract Is Nothing _
    Then
        Set tickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)
        Set tickerConfigSection = tickersConfigSection.GetConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
        If Not tickerConfigSection Is Nothing Then tickerConfigSection.MoveUp
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="moveRowUpInConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
                
End Sub

Private Sub removeTicker( _
                ByVal index As Long)
Dim gridRowIndex As Long
Dim i As Long
Dim rowdata As Long
Dim lTicker As Ticker

Dim failpoint As Long
On Error GoTo Err

gridRowIndex = mTickerTable(index).tickerGridRow

TickerGrid.RemoveItem gridRowIndex
mNextGridRowIndex = mNextGridRowIndex - 1

Set lTicker = mTickerTable(index).theTicker
removeTickerFromConfig lTicker

Set mTickerTable(index).theTicker = Nothing
mTickerTable(index).tickerGridRow = 0

For i = gridRowIndex To mNextGridRowIndex - 1
    rowdata = TickerGrid.rowdata(i)
    mTickerTable(rowdata).tickerGridRow = i
Next

mSelectedTickers.Remove lTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="removeTicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub removeTickerFromConfig( _
                ByVal pTicker As Ticker)
Dim tickersConfigSection As ConfigurationSection
Dim tickerConfigSection As ConfigurationSection

Dim failpoint As Long
On Error GoTo Err

If Not mConfig Is Nothing Then
    If Not pTicker.ReplayingTickfile And _
        Not pTicker.Contract Is Nothing _
    Then
        Set tickersConfigSection = mConfig.AddPrivateConfigurationSection(ConfigSectionTickers)
        Set tickerConfigSection = tickersConfigSection.GetConfigurationSection(ConfigSectionTicker & "(" & pTicker.Key & ")")
        If Not tickerConfigSection Is Nothing Then tickersConfigSection.RemoveConfigurationSection ConfigSectionTicker & "(" & pTicker.Key & ")"
    End If
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="removeTickerFromConfig", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub selectATicker( _
                ByVal row As Long)
Dim index As Long
Dim failpoint As Long
On Error GoTo Err

index = TickerGrid.rowdata(row)

If Not mTickerTable(index).theTicker Is Nothing Then
    mSelectedTickers.Add mTickerTable(index).theTicker
    highlightRow row
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="selectATicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setColumnWidth( _
                ByVal col As Long, _
                ByVal widthChars As Long, _
                ByVal isLetters As Boolean)
Dim failpoint As Long
On Error GoTo Err

TickerGrid.colWidth(mColumnMap(col)) = IIf(isLetters, mLetterWidth, mDigitWidth) * widthChars

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setColumnWidth", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setColumnWidths()
Dim failpoint As Long
On Error GoTo Err

TickerGrid.Redraw = False
setColumnWidth TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, True
setColumnWidth TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, True
setColumnWidth TickerGridColumns.currencyCode, TickerGridColumnWidths.CurrencyWidth, True
setColumnWidth TickerGridColumns.BidSize, TickerGridColumnWidths.BidSizeWidth, False
setColumnWidth TickerGridColumns.Bid, TickerGridColumnWidths.BidWidth, False
setColumnWidth TickerGridColumns.Ask, TickerGridColumnWidths.AskWidth, False
setColumnWidth TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, False
setColumnWidth TickerGridColumns.Trade, TickerGridColumnWidths.TradeWidth, False
setColumnWidth TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSizeWidth, False
setColumnWidth TickerGridColumns.Volume, TickerGridColumnWidths.VolumeWidth, False
setColumnWidth TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, False
setColumnWidth TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, False
setColumnWidth TickerGridColumns.HighPrice, TickerGridColumnWidths.highWidth, False
setColumnWidth TickerGridColumns.LowPrice, TickerGridColumnWidths.LowWidth, False
setColumnWidth TickerGridColumns.OpenPrice, TickerGridColumnWidths.OpenWidth, False
setColumnWidth TickerGridColumns.ClosePrice, TickerGridColumnWidths.CloseWidth, False
setColumnWidth TickerGridColumns.OpenInterest, TickerGridColumnWidths.openInterestWidth, False
setColumnWidth TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, True
setColumnWidth TickerGridColumns.symbol, TickerGridColumnWidths.SymbolWidth, True
setColumnWidth TickerGridColumns.secType, TickerGridColumnWidths.SecTypeWidth, True
setColumnWidth TickerGridColumns.expiry, TickerGridColumnWidths.ExpiryWidth, True
setColumnWidth TickerGridColumns.exchange, TickerGridColumnWidths.ExchangeWidth, True
setColumnWidth TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, True
setColumnWidth TickerGridColumns.strike, TickerGridColumnWidths.StrikeWidth, False
TickerGrid.Redraw = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setColumnWidths", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setupColumnMap( _
                    ByVal maxIndex As Long)
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

ReDim mColumnMap(maxIndex) As Long
For i = 0 To UBound(mColumnMap)
    mColumnMap(i) = i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupColumnMap", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setupDefaultTickerGridColumns()

Dim failpoint As Long
On Error GoTo Err

setupColumnMap TickerGridColumns.MaxColumn

setupTickerGridColumn TickerGridColumns.Selector, TickerGridColumnWidths.SelectorWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.TickerName, TickerGridColumnWidths.NameWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.currencyCode, TickerGridColumnWidths.CurrencyWidth, True, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.BidSize, TickerGridColumnWidths.BidSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Bid, TickerGridColumnWidths.BidWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Ask, TickerGridColumnWidths.AskWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.AskSize, TickerGridColumnWidths.AskSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Trade, TickerGridColumnWidths.TradeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.TradeSize, TickerGridColumnWidths.TradeSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Volume, TickerGridColumnWidths.VolumeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Change, TickerGridColumnWidths.ChangeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.ChangePercent, TickerGridColumnWidths.ChangePercentWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.HighPrice, TickerGridColumnWidths.highWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.LowPrice, TickerGridColumnWidths.LowWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.OpenPrice, TickerGridColumnWidths.OpenWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.ClosePrice, TickerGridColumnWidths.CloseWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.OpenInterest, TickerGridColumnWidths.openInterestWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
setupTickerGridColumn TickerGridColumns.Description, TickerGridColumnWidths.DescriptionWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.symbol, TickerGridColumnWidths.SymbolWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.secType, TickerGridColumnWidths.SecTypeWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.expiry, TickerGridColumnWidths.ExpiryWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.exchange, TickerGridColumnWidths.ExchangeWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.OptionRight, TickerGridColumnWidths.OptionRightWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridColumns.strike, TickerGridColumnWidths.StrikeWidth, False, TWControls10.AlignmentSettings.TwGridAlignCenterCenter

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupDefaultTickerGridColumns", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setupDefaultTickerGridHeaders()

Dim failpoint As Long
On Error GoTo Err

TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Selector)) = ""
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.TickerName)) = "Name"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.currencyCode)) = "Curr"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.BidSize)) = "Bid size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Bid)) = "Bid"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Ask)) = "Ask"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.AskSize)) = "Ask size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Trade)) = "Last"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.TradeSize)) = "Last size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Volume)) = "Volume"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Change)) = "Chg"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.ChangePercent)) = "Chg %"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.HighPrice)) = "High"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.LowPrice)) = "Low"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.OpenPrice)) = "Open"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.ClosePrice)) = "Close"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.OpenInterest)) = "Open interest"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.Description)) = "Description"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.symbol)) = "Symbol"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.secType)) = "Sec Type"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.expiry)) = "Expiry"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.exchange)) = "Exchange"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.OptionRight)) = "Right"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridColumns.strike)) = "Strike"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupDefaultTickerGridHeaders", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setupSummaryTickerGridColumns()
Dim failpoint As Long
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

setupTickerGridColumn TickerGridSummaryColumns.Selector, TickerGridSummaryColumnWidths.SelectorWidth, True, TWControls10.AlignmentSettings.TwGridAlignCenterBottom
setupTickerGridColumn TickerGridSummaryColumns.TickerName, TickerGridSummaryColumnWidths.NameWidth, True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.BidSize, TickerGridSummaryColumnWidths.BidSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Bid, TickerGridSummaryColumnWidths.BidWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Ask, TickerGridSummaryColumnWidths.AskWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.AskSize, TickerGridSummaryColumnWidths.AskSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Trade, TickerGridSummaryColumnWidths.TradeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.TradeSize, TickerGridSummaryColumnWidths.TradeSizeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Volume, TickerGridSummaryColumnWidths.VolumeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.Change, TickerGridSummaryColumnWidths.ChangeWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.ChangePercent, TickerGridSummaryColumnWidths.ChangePercentWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupTickerGridColumn TickerGridSummaryColumns.OpenInterest, TickerGridSummaryColumnWidths.openInterestWidth, False, TWControls10.AlignmentSettings.TwGridAlignLeftCenter

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupSummaryTickerGridColumns", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setupSummaryTickerGridHeaders()

Dim failpoint As Long
On Error GoTo Err

TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Selector)) = ""
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.TickerName)) = "Name"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.BidSize)) = "Bid size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Bid)) = "Bid"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Ask)) = "Ask"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.AskSize)) = "Ask size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Trade)) = "Last"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.TradeSize)) = "Last size"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Volume)) = "Volume"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.Change)) = "Chg"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.ChangePercent)) = "Chg %"
TickerGrid.TextMatrix(0, mColumnMap(TickerGridSummaryColumns.OpenInterest)) = "Open interest"

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupSummaryTickerGridHeaders", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Sub

Private Sub setupTickerGridColumn( _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Long, _
                ByVal isLetters As Boolean, _
                ByVal align As TWControls10.AlignmentSettings)
    
Dim failpoint As Long
On Error GoTo Err

With TickerGrid
    
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .colWidth(columnNumber) = 0
    End If
    
    .ColData(columnNumber) = columnNumber
    
    setColumnWidth columnNumber, columnWidth, isLetters
    
    .ColAlignment(columnNumber) = align
    .ColAlignmentFixed(columnNumber) = TwGridAlignCenterCenter
End With

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="setupTickerGridColumn", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub stopTicker( _
                ByVal index As Long)
Dim lTicker As Ticker

Dim failpoint As Long
On Error GoTo Err

Set lTicker = mTickerTable(index).theTicker
lTicker.RemoveQuoteListener Me
lTicker.RemovePriceChangeListener Me

removeTicker index

lTicker.stopTicker

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="stopTicker", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub storeColumnMap()
Dim i As Long
Dim s As String

Dim failpoint As Long
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

For i = 0 To UBound(mColumnMap)
    s = s & IIf(s = "", "", ", ") & CStr(mColumnMap(i))
Next

mConfig.SetSetting ConfigSettingColumnMap, s

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="storeColumnMap", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Private Sub toggleRowSelection( _
                ByVal row As Long)
Dim failpoint As Long
On Error GoTo Err

If isRowSelected(row) Then
    deselectATicker row
Else
    selectATicker row
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:="toggleRowSelection", pNumber:=Err.Number, pSource:=Err.Source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

