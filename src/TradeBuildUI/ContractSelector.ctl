VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#37.0#0"; "TWControls10.ocx"
Begin VB.UserControl ContractSelector 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin TWControls10.TWGrid TWGrid1 
      Height          =   3255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   5741
   End
End
Attribute VB_Name = "ContractSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IContractSelector

'@================================================================================
' Events
'@================================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event DblClick()
Attribute DblClick.VB_UserMemId = -601

'@================================================================================
' Enums
'@================================================================================

Private Enum ContractsGridColumns
    secType
    localSymbol = secType
    exchange
    expiry = exchange
    currencyCode
    strike = currencyCode
    OptionRight
'    Filler
'    Description
    MaxColumn = OptionRight
End Enum

' Character widths of the twgrid1 columns
Private Enum ContractsGridColumnWidths
'    LocalSymbolWidth = 15
    SecTypeWidth = 15
    ExchangeWidth = 10
'    ExpiryWidth = 10
    CurrencyWidth = 9
    'StrikeWidth = 9
    OptionRightWidth = 5
    DescriptionWidth = 20
    FillerWidth = 500
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractSelector"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                                  As Contracts

Private mIncludeHistoricalContracts                 As Boolean

Private mLetterWidth                                As Single
Private mDigitWidth                                 As Single

Private mSortKeys()                                 As ContractSortKeyIds

Private mCurrSectype                                As SecurityTypes
Private mCurrCurrency                               As String
Private mCurrExchange                               As String

Private mControlDown                                As Boolean
Private mShiftDown                                  As Boolean
Private mAltDown                                    As Boolean

Private mSelectedRows                               As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_GotFocus()
TWGrid1.SetFocus
End Sub

Private Sub UserControl_Initialize()
Dim widthString As String
widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = UserControl.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = UserControl.TextWidth(widthString) / Len(widthString)

ReDim mSortKeys(5) As ContractSortKeyIds
mSortKeys(0) = ContractSortKeySecType
mSortKeys(1) = ContractSortKeyExchange
mSortKeys(2) = ContractSortKeyCurrency
mSortKeys(3) = ContractSortKeyExpiry
mSortKeys(4) = ContractSortKeyStrike
mSortKeys(5) = ContractSortKeyRight

End Sub

Private Sub UserControl_InitProperties()
RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
ScrollBars = flexScrollBarBoth
setupGrid

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

TWGrid1.RowBackColorOdd = PropBag.ReadProperty("RowBackColorOdd", CRowBackColorOdd)
TWGrid1.RowBackColorEven = PropBag.ReadProperty("RowBackColorEven", CRowBackColorEven)
TWGrid1.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)

setupGrid

End Sub

Private Sub UserControl_Resize()
TWGrid1.Move 0, 0, UserControl.Width, UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("ScrollBars", TWGrid1.ScrollBars, 3)
    Call PropBag.WriteProperty("RowBackColorOdd", TWGrid1.RowBackColorOdd, 0)
    Call PropBag.WriteProperty("RowBackColorEven", TWGrid1.RowBackColorEven, 0)
End Sub

'@================================================================================
' IContractSelector Interface Members
'@================================================================================

Private Sub IContractSelector_Initialise(ByVal pContracts As ContractUtils26.Contracts)
Initialise pContracts
End Sub

Private Property Get IContractSelector_SelectedContracts() As ContractUtils26.Contracts
Set IContractSelector_SelectedContracts = SelectedContracts
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TWGrid1_Click()
Dim row As Long
Dim rowSel As Long
Dim col As Long
Dim colSel As Long
row = TWGrid1.row
rowSel = TWGrid1.rowSel
col = TWGrid1.col
colSel = TWGrid1.colSel

If TWGrid1.rowdata(row) = 0 Then Exit Sub

If Not mControlDown Then
    deselectSelectedContracts
    selectContract row
Else
    toggleRowSelection row
End If

RaiseEvent Click
End Sub

Private Sub TWGrid1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub TWGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub TWGrid1_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub TWGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub TWGrid1_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)

RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub TWGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub TWGrid1_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)

RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let IncludeHistoricalContracts( _
                ByVal value As Boolean)
mIncludeHistoricalContracts = value
End Property

Public Property Get IncludeHistoricalContracts() As Boolean
IncludeHistoricalContracts = mIncludeHistoricalContracts
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
    RowBackColorOdd = TWGrid1.RowBackColorOdd
End Property

Public Property Let RowBackColorOdd(ByVal New_RowBackColorOdd As OLE_COLOR)
    TWGrid1.RowBackColorOdd = New_RowBackColorOdd
    PropertyChanged "RowBackColorOdd"
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
    RowBackColorEven = TWGrid1.RowBackColorEven
End Property

Public Property Let RowBackColorEven(ByVal New_RowBackColorEven As OLE_COLOR)
    TWGrid1.RowBackColorEven = New_RowBackColorEven
    PropertyChanged "RowBackColorEven"
End Property

Public Property Get ScrollBars() As ScrollBarsSettings
    ScrollBars = TWGrid1.ScrollBars
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
    TWGrid1.ScrollBars = New_ScrollBars
    PropertyChanged "ScrollBars"
End Property

Public Property Get SelectedContracts() As Contracts
Dim scb As ContractsBuilder
Dim i As Long
Dim row As Long

Set scb = CreateContractsBuilder(mContracts.contractSpecifier)

For i = 1 To mSelectedRows.Count
    row = mSelectedRows(i)
    scb.AddContract mContracts(TWGrid1.rowdata(row))
Next

Set SelectedContracts = scb.Contracts
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pContracts As Contracts)
Dim lRow As Long
Dim lContract As Contract
Dim contractSpec As contractSpecifier
Dim index As Long

TWGrid1.ClearStructure
setupGrid

TWGrid1.Redraw = False

Set mContracts = pContracts
mContracts.SortKeys = mSortKeys
mContracts.Sort

Set mSelectedRows = New Collection

lRow = -1

For Each lContract In mContracts
    index = index + 1
    Set contractSpec = lContract.specifier
    
    If IncludeHistoricalContracts Or Not isExpired(lContract) Then
        lRow = lRow + 1
        If lRow > TWGrid1.Rows - 1 Then TWGrid1.Rows = TWGrid1.Rows + 1
        
        TWGrid1.row = lRow
        
        If needHeadingRow(contractSpec) Then
            writeHeadingRow contractSpec
            lRow = lRow + 1
            If lRow > TWGrid1.Rows - 1 Then TWGrid1.Rows = TWGrid1.Rows + 1
            TWGrid1.row = lRow
        End If
        
        TWGrid1.rowdata(lRow) = index
        
        writeRow lContract, contractSpec
        
        mCurrSectype = contractSpec.secType
        mCurrCurrency = contractSpec.currencyCode
        mCurrExchange = contractSpec.exchange
    End If
Next

TWGrid1.Redraw = True

mCurrSectype = SecTypeNone
mCurrCurrency = ""
mCurrExchange = ""
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub deselectContract( _
                ByVal row As Long)
mSelectedRows.Remove CStr(row)
highlightRow row
End Sub

Private Sub deselectSelectedContracts()
Dim i As Long
Dim row As Long

For i = mSelectedRows.Count To 1 Step -1
    row = mSelectedRows(i)
    deselectContract row
Next
End Sub

Private Sub highlightRow(ByVal rowNumber As Long)
Dim i As Long

If rowNumber < 0 Then Exit Sub

TWGrid1.row = rowNumber

For i = 1 To TWGrid1.Cols - 1
    TWGrid1.col = i
    If TWGrid1.CellFontBold Then
        TWGrid1.CellFontBold = False
    Else
        TWGrid1.CellFontBold = True
    End If
Next

TWGrid1.col = 0
TWGrid1.colSel = ContractsGridColumns.MaxColumn
TWGrid1.InvertCellColors

End Sub

Private Function isExpired( _
                ByVal pContract As Contract) As Boolean
If pContract.specifier.secType = SecTypeFuture Or _
    pContract.specifier.secType = SecTypeFuturesOption Or _
    pContract.specifier.secType = SecTypeOption _
Then
    If Int(pContract.expiryDate) < Int(Now) Then isExpired = True
End If
End Function

Private Function isFullHeadingSecType( _
                ByVal secType As SecurityTypes) As Boolean
If secType = SecTypeFuture Or _
    secType = SecTypeFuturesOption Or _
    secType = SecTypeOption _
Then
    isFullHeadingSecType = True
End If
End Function

Private Function isHeadingWithoutExchangeSecType( _
                ByVal secType As SecurityTypes)
If secType = SecTypeStock Or _
    secType = SecTypeIndex _
Then
    isHeadingWithoutExchangeSecType = True
End If
End Function

Private Function isHeadingWithoutCurrencySecType( _
                ByVal secType As SecurityTypes)
If secType = SecTypeStock Or _
    secType = SecTypeCash Or _
    secType = SecTypeIndex _
Then
    isHeadingWithoutCurrencySecType = True
End If
End Function

Private Function isRowSelected( _
                ByVal row As Long) As Boolean
On Error Resume Next
isRowSelected = (CLng(mSelectedRows(CStr(row))) = row)
End Function

Private Function needFullHeadingRow( _
                ByVal contractSpec As contractSpecifier) As Boolean
If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.currencyCode <> mCurrCurrency Or _
    contractSpec.exchange <> mCurrExchange) And _
    isFullHeadingSecType(contractSpec.secType) _
Then
    needFullHeadingRow = True
End If
End Function

Private Function needHeadingRow( _
                ByVal contractSpec As contractSpecifier) As Boolean
If needFullHeadingRow(contractSpec) Or _
    needHeadingRowWithoutExchange(contractSpec) Or _
    needHeadingRowWithoutCurrency(contractSpec) Or _
    needHeadingRowWithSectypeOnly(contractSpec) _
Then
    needHeadingRow = True
End If
End Function

Private Function needHeadingRowWithoutExchange( _
                ByVal contractSpec As contractSpecifier) As Boolean
If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.currencyCode <> mCurrCurrency) And _
    isHeadingWithoutExchangeSecType(contractSpec.secType) And _
    (Not isHeadingWithoutExchangeSecType(contractSpec.secType)) _
Then
    needHeadingRowWithoutExchange = True
End If
End Function

Private Function needHeadingRowWithoutCurrency( _
                ByVal contractSpec As contractSpecifier) As Boolean
If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.exchange <> mCurrExchange) And _
    isHeadingWithoutCurrencySecType(contractSpec.secType) And _
    (Not isHeadingWithoutExchangeSecType(contractSpec.secType)) _
Then
    needHeadingRowWithoutCurrency = True
End If
End Function

Private Function needHeadingRowWithSectypeOnly( _
                ByVal contractSpec As contractSpecifier) As Boolean
If contractSpec.secType <> mCurrSectype And _
    isHeadingWithoutExchangeSecType(contractSpec.secType) And _
    isHeadingWithoutCurrencySecType(contractSpec.secType) _
Then
    needHeadingRowWithSectypeOnly = True
End If
End Function

Private Sub selectContract( _
                ByVal row As Long)
mSelectedRows.Add row, CStr(row)
highlightRow row
End Sub

Private Sub setupGrid()
'TWGrid1.AllowBigSelection = True
TWGrid1.Cols = 2
TWGrid1.GridLineWidth = 0
TWGrid1.FillStyle = TwGridFillRepeat
TWGrid1.FixedRows = 0
TWGrid1.FixedCols = 0
TWGrid1.HighLight = TwGridHighlightNever
TWGrid1.Rows = 1
TWGrid1.SelectionMode = TwGridSelectionByRow
TWGrid1.BackColorBkg = SystemColorConstants.vbWindowBackground
TWGrid1.FocusRect = TwGridFocusNone

'setupGridColumn 0, ContractsGridColumns.localSymbol, ContractsGridColumnWidths.LocalSymbolWidth, "Symbol", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupGridColumn 0, ContractsGridColumns.secType, ContractsGridColumnWidths.SecTypeWidth, "Sec Type", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupGridColumn 0, ContractsGridColumns.exchange, ContractsGridColumnWidths.ExchangeWidth, "Exchange", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
'setupGridColumn 0, ContractsGridColumns.expiry, ContractsGridColumnWidths.ExpiryWidth, "Expiry", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
setupGridColumn 0, ContractsGridColumns.currencyCode, ContractsGridColumnWidths.CurrencyWidth, "Curr", True, TWControls10.AlignmentSettings.TwGridAlignCenterCenter
'setupGridColumn 0, ContractsGridColumns.strike, ContractsGridColumnWidths.StrikeWidth, "Strike", False, TWControls10.AlignmentSettings.TwGridAlignRightCenter
setupGridColumn 0, ContractsGridColumns.OptionRight, ContractsGridColumnWidths.OptionRightWidth, "Right", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
'setupGridColumn 0, ContractsGridColumns.Description, ContractsGridColumnWidths.DescriptionWidth, "Description", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
'setupGridColumn 0, ContractsGridColumns.Filler, ContractsGridColumnWidths.FillerWidth, "", True, TWControls10.AlignmentSettings.TwGridAlignLeftCenter
End Sub

Private Sub setupGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal columnHeader As String, _
                ByVal isLetters As Boolean, _
                ByVal align As TWControls10.AlignmentSettings)
    
Dim lColumnWidth As Long

With TWGrid1
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
'    .ColAlignmentFixed(columnNumber) = align
'    .TextMatrix(rowNumber, columnNumber) = columnHeader
End With
End Sub

Private Sub toggleRowSelection( _
                ByVal row As Long)
If isRowSelected(row) Then
    deselectContract row
Else
    selectContract row
End If
End Sub

Private Sub writeHeadingRow( _
                ByVal contractSpec As contractSpecifier)
Dim excludeExchange As Boolean
Dim excludeCurrency As Boolean

If isHeadingWithoutExchangeSecType(contractSpec.secType) Then excludeExchange = True
If isHeadingWithoutCurrencySecType(contractSpec.secType) Then excludeCurrency = True

TWGrid1.col = 0
TWGrid1.colSel = ContractsGridColumns.MaxColumn
Select Case contractSpec.secType
    Case SecTypeStock
        TWGrid1.CellBackColor = &H359F51
    Case SecTypeFuture
        TWGrid1.CellBackColor = &H345DA0
    Case SecTypeOption
        TWGrid1.CellBackColor = &HB4B650
    Case SecTypeFuturesOption
        TWGrid1.CellBackColor = &H8B5BAB
    Case SecTypeCash
        TWGrid1.CellBackColor = &H989044
    Case SecTypeCombo
        TWGrid1.CellBackColor = &H6A7E98
    Case SecTypeIndex
        TWGrid1.CellBackColor = &HC84A50
End Select
TWGrid1.CellFontBold = True
TWGrid1.CellForeColor = vbWhite

TWGrid1.col = ContractsGridColumns.secType
TWGrid1.Text = SecTypeToString(contractSpec.secType)
        
If Not excludeCurrency Then
    TWGrid1.col = ContractsGridColumns.currencyCode
    TWGrid1.Text = contractSpec.currencyCode
End If

If Not excludeExchange Then
    TWGrid1.col = ContractsGridColumns.exchange
    TWGrid1.Text = contractSpec.exchange
End If
End Sub

Private Sub writeRow( _
                ByVal pContract As Contract, _
                ByVal contractSpec As contractSpecifier)
TWGrid1.col = 0
TWGrid1.colSel = ContractsGridColumns.MaxColumn
TWGrid1.CellBackColor = vbBlack ' Clear out any cell background colour
TWGrid1.CellFontBold = False
TWGrid1.CellForeColor = vbBlack

TWGrid1.col = ContractsGridColumns.localSymbol
TWGrid1.Text = contractSpec.localSymbol

If isFullHeadingSecType(contractSpec.secType) Then
Else
    If isHeadingWithoutExchangeSecType(contractSpec.secType) Then
        TWGrid1.col = ContractsGridColumns.exchange
        TWGrid1.Text = contractSpec.exchange
    End If
    If isHeadingWithoutCurrencySecType(contractSpec.secType) Then
        TWGrid1.col = ContractsGridColumns.currencyCode
        TWGrid1.Text = contractSpec.currencyCode
    End If
End If
    
'TWGrid1.col = ContractsGridColumns.Description
'TWGrid1.Text = lContract.Description

Select Case contractSpec.secType
    Case SecTypeFuture
        TWGrid1.col = ContractsGridColumns.expiry
        TWGrid1.Text = FormatDateTime(pContract.expiryDate, vbShortDate)
    Case SecTypeOption, SecTypeFuturesOption
        TWGrid1.col = ContractsGridColumns.expiry
        TWGrid1.Text = FormatDateTime(pContract.expiryDate, vbShortDate)
    
        TWGrid1.col = ContractsGridColumns.OptionRight
        TWGrid1.Text = OptionRightToString(contractSpec.Right)
        
        TWGrid1.col = ContractsGridColumns.strike
        TWGrid1.Text = Format(contractSpec.strike, pContract.PriceFormatString)
        TWGrid1.CellAlignment = TwGridAlignRightCenter
    'Case SecTypeCombo

End Select

End Sub
