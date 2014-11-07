VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#16.1#0"; "TWControls40.ocx"
Begin VB.UserControl ContractSelector 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin TWControls40.TWGrid TWGrid1 
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
    LocalSymbol = secType
    Exchange
    Expiry = Exchange
    CurrencyCode
    Strike = CurrencyCode
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

Private mContracts                                  As IContracts
Private mAllowMultipleSelection                     As Boolean

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

Private mCount                                      As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_GotFocus()
Const ProcName As String = "UserControl_GotFocus"
On Error GoTo Err

TWGrid1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

Dim widthString As String
widthString = "ABCDEFGH IJKLMNOP QRST UVWX YZ"
mLetterWidth = UserControl.TextWidth(widthString) / Len(widthString)
widthString = ".0123456789"
mDigitWidth = UserControl.TextWidth(widthString) / Len(widthString)

ReDim mSortKeys(6) As ContractSortKeyIds
mSortKeys(0) = ContractSortKeySecType
mSortKeys(1) = ContractSortKeySymbol
mSortKeys(2) = ContractSortKeyExchange
mSortKeys(3) = ContractSortKeyCurrency
mSortKeys(4) = ContractSortKeyExpiry
mSortKeys(5) = ContractSortKeyStrike
mSortKeys(6) = ContractSortKeyRight

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_InitProperties()
Const ProcName As String = "UserControl_InitProperties"
On Error GoTo Err

RowBackColorEven = CRowBackColorEven
RowBackColorOdd = CRowBackColorOdd
ScrollBars = flexScrollBarBoth
setupGrid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

TWGrid1.RowBackColorOdd = PropBag.ReadProperty("RowBackColorOdd", CRowBackColorOdd)
TWGrid1.RowBackColorEven = PropBag.ReadProperty("RowBackColorEven", CRowBackColorEven)
TWGrid1.ScrollBars = PropBag.ReadProperty("ScrollBars", 3)

setupGrid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

If UserControl.Width = 0 Or UserControl.Height = 0 Then Exit Sub

TWGrid1.Move 0, 0, UserControl.Width, UserControl.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

On Error Resume Next
Call PropBag.WriteProperty("ScrollBars", TWGrid1.ScrollBars, 3)
Call PropBag.WriteProperty("RowBackColorOdd", TWGrid1.RowBackColorOdd, 0)
Call PropBag.WriteProperty("RowBackColorEven", TWGrid1.RowBackColorEven, 0)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' IContractSelector Interface Members
'@================================================================================

Private Sub IContractSelector_Initialise(ByVal pContracts As IContracts, ByVal pAllowMultipleSelection As Boolean)
Const ProcName As String = "IContractSelector_Initialise"
On Error GoTo Err

Initialise pContracts, pAllowMultipleSelection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IContractSelector_SelectedContracts() As IContracts
Const ProcName As String = "IContractSelector_SelectedContracts"
On Error GoTo Err

Set IContractSelector_SelectedContracts = SelectedContracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TWGrid1_Click()
Const ProcName As String = "TWGrid1_Click"
On Error GoTo Err

Dim lRow As Long
lRow = TWGrid1.Row

Dim lRowSel As Long
lRowSel = TWGrid1.RowSel

Dim lCol As Long
lCol = TWGrid1.col

Dim lColSel As Long
lColSel = TWGrid1.ColSel

If TWGrid1.RowData(lRow) = 0 Then Exit Sub

If Not mControlDown Or Not mAllowMultipleSelection Then
    deselectSelectedContracts
    selectContract lRow
Else
    toggleRowSelection lRow
End If

RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_DblClick()
Const ProcName As String = "TWGrid1_DblClick"
On Error GoTo Err

RaiseEvent DblClick

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TWGrid1_KeyDown"
On Error GoTo Err

RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_KeyPress(KeyAscii As Integer)
Const ProcName As String = "TWGrid1_KeyPress"
On Error GoTo Err

RaiseEvent KeyPress(KeyAscii)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "TWGrid1_KeyUp"
On Error GoTo Err

RaiseEvent KeyUp(KeyCode, Shift)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "TWGrid1_MouseDown"
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)

RaiseEvent MouseDown(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "TWGrid1_MouseMove"
On Error GoTo Err

RaiseEvent MouseMove(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TWGrid1_MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "TWGrid1_MouseUp"
On Error GoTo Err

mShiftDown = (Shift And KeyDownShift)
mControlDown = (Shift And KeyDownCtrl)
mAltDown = (Shift And KeyDownAlt)

RaiseEvent MouseUp(Button, Shift, X, Y)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mCount
End Property

Public Property Let IncludeHistoricalContracts( _
                ByVal value As Boolean)
mIncludeHistoricalContracts = value
End Property

Public Property Get IncludeHistoricalContracts() As Boolean
IncludeHistoricalContracts = mIncludeHistoricalContracts
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

RowBackColorEven = TWGrid1.RowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorEven(ByVal New_RowBackColorEven As OLE_COLOR)
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

TWGrid1.RowBackColorEven = New_RowBackColorEven
PropertyChanged "RowBackColorEven"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

RowBackColorOdd = TWGrid1.RowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorOdd(ByVal New_RowBackColorOdd As OLE_COLOR)
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

TWGrid1.RowBackColorOdd = New_RowBackColorOdd
PropertyChanged "RowBackColorOdd"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ScrollBars() As ScrollBarsSettings
Const ProcName As String = "ScrollBars"
On Error GoTo Err

ScrollBars = TWGrid1.ScrollBars

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScrollBars(ByVal New_ScrollBars As ScrollBarsSettings)
Const ProcName As String = "ScrollBars"
On Error GoTo Err

TWGrid1.ScrollBars = New_ScrollBars
PropertyChanged "ScrollBars"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedContracts() As IContracts
Const ProcName As String = "SelectedContracts"
On Error GoTo Err

Dim scb As IContractsBuilder
Set scb = New ContractsBuilder

Dim i As Long
For i = 1 To mSelectedRows.Count
    Dim Row As Long
    Row = mSelectedRows(i)
    scb.Add mContracts.ItemAtIndex(TWGrid1.RowData(Row))
Next

Set SelectedContracts = scb.Contracts

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TextboxBackColor(ByVal value As OLE_COLOR)
Const ProcName As String = "TextboxBackColor"
On Error GoTo Err

TWGrid1.RowBackColorEven = value
TWGrid1.RowBackColorOdd = toneDown(value)
TWGrid1.BackColorBkg = value


Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextboxBackColor() As OLE_COLOR
TextboxBackColor = TWGrid1.RowBackColorEven
End Property

Public Property Let TextboxForeColor(ByVal value As OLE_COLOR)
Const ProcName As String = "TextboxForeColor"
On Error GoTo Err

TWGrid1.ForeColor = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextboxForeColor() As OLE_COLOR
TextboxForeColor = TWGrid1.ForeColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pContracts As IContracts, _
                ByVal pAllowMultipleSelection As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

mAllowMultipleSelection = pAllowMultipleSelection

TWGrid1.Clear

TWGrid1.Redraw = False

Set mContracts = pContracts
mContracts.SortKeys = mSortKeys

Dim et As New ElapsedTimer
et.StartTiming
gLogger.Log "Sorted contracts: elapsed time (millisecs):", ProcName, ModuleName, LogLevelDetail, Int(et.ElapsedTimeMicroseconds / 1000#)
Set mSelectedRows = New Collection

Dim lRow As Long
lRow = -1

Dim lIndex As Long
Dim lContract As Contract
For Each lContract In mContracts
    lIndex = lIndex + 1
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = lContract.Specifier
    
    If IncludeHistoricalContracts Or Not IsContractExpired(lContract) Then
        lRow = lRow + 1
        If lRow > TWGrid1.Rows - 1 Then TWGrid1.Rows = TWGrid1.Rows + 1
        
        TWGrid1.Row = lRow
        
        If needHeadingRow(lContractSpec) Then
            writeHeadingRow lContractSpec
            lRow = lRow + 1
            If lRow > TWGrid1.Rows - 1 Then TWGrid1.Rows = TWGrid1.Rows + 1
            TWGrid1.Row = lRow
        End If
        
        TWGrid1.RowData(lRow) = lIndex
        
        writeRow lContract
        
        mCurrSectype = lContractSpec.secType
        mCurrCurrency = lContractSpec.CurrencyCode
        mCurrExchange = lContractSpec.Exchange
    End If
Next

TWGrid1.Redraw = True

mCurrSectype = SecTypeNone
mCurrCurrency = ""
mCurrExchange = ""

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub deselectContract( _
                ByVal Row As Long)
Const ProcName As String = "deselectContract"
On Error GoTo Err

mSelectedRows.Remove CStr(Row)
highlightRow Row

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deselectSelectedContracts()
Const ProcName As String = "deselectSelectedContracts"
On Error GoTo Err

Dim i As Long
For i = mSelectedRows.Count To 1 Step -1
    Dim Row As Long
    Row = mSelectedRows(i)
    deselectContract Row
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub highlightRow(ByVal rowNumber As Long)
Const ProcName As String = "highlightRow"
On Error GoTo Err

If rowNumber < 0 Then Exit Sub

TWGrid1.Row = rowNumber

Dim i As Long
For i = 1 To TWGrid1.Cols - 1
    TWGrid1.col = i
    If TWGrid1.CellFontBold Then
        TWGrid1.CellFontBold = False
    Else
        TWGrid1.CellFontBold = True
    End If
Next

TWGrid1.col = 0
TWGrid1.ColSel = ContractsGridColumns.MaxColumn
TWGrid1.InvertCellColors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function isFullHeadingSecType( _
                ByVal secType As SecurityTypes) As Boolean
Const ProcName As String = "isFullHeadingSecType"
On Error GoTo Err

If secType = SecTypeFuture Or _
    secType = SecTypeFuturesOption Or _
    secType = SecTypeOption _
Then
    isFullHeadingSecType = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isHeadingWithoutExchangeSecType( _
                ByVal secType As SecurityTypes)
Const ProcName As String = "isHeadingWithoutExchangeSecType"
On Error GoTo Err

If secType = SecTypeStock Or _
    secType = SecTypeIndex _
Then
    isHeadingWithoutExchangeSecType = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isHeadingWithoutCurrencySecType( _
                ByVal secType As SecurityTypes)
Const ProcName As String = "isHeadingWithoutCurrencySecType"
On Error GoTo Err

If secType = SecTypeStock Or _
    secType = SecTypeCash Or _
    secType = SecTypeIndex _
Then
    isHeadingWithoutCurrencySecType = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isRowSelected( _
                ByVal Row As Long) As Boolean
Const ProcName As String = "isRowSelected"
On Error GoTo Err

On Error Resume Next
isRowSelected = (CLng(mSelectedRows(CStr(Row))) = Row)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needFullHeadingRow( _
                ByVal contractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "needFullHeadingRow"
On Error GoTo Err

If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.CurrencyCode <> mCurrCurrency Or _
    contractSpec.Exchange <> mCurrExchange) And _
    isFullHeadingSecType(contractSpec.secType) _
Then
    needFullHeadingRow = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needHeadingRow( _
                ByVal contractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "needHeadingRow"
On Error GoTo Err

If needFullHeadingRow(contractSpec) Or _
    needHeadingRowWithoutExchange(contractSpec) Or _
    needHeadingRowWithoutCurrency(contractSpec) Or _
    needHeadingRowWithSectypeOnly(contractSpec) _
Then
    needHeadingRow = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needHeadingRowWithoutExchange( _
                ByVal contractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "needHeadingRowWithoutExchange"
On Error GoTo Err

If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.CurrencyCode <> mCurrCurrency) And _
    isHeadingWithoutExchangeSecType(contractSpec.secType) And _
    (Not isHeadingWithoutExchangeSecType(contractSpec.secType)) _
Then
    needHeadingRowWithoutExchange = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needHeadingRowWithoutCurrency( _
                ByVal contractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "needHeadingRowWithoutCurrency"
On Error GoTo Err

If (contractSpec.secType <> mCurrSectype Or _
    contractSpec.Exchange <> mCurrExchange) And _
    isHeadingWithoutCurrencySecType(contractSpec.secType) And _
    (Not isHeadingWithoutExchangeSecType(contractSpec.secType)) _
Then
    needHeadingRowWithoutCurrency = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needHeadingRowWithSectypeOnly( _
                ByVal contractSpec As IContractSpecifier) As Boolean
Const ProcName As String = "needHeadingRowWithSectypeOnly"
On Error GoTo Err

If contractSpec.secType <> mCurrSectype And _
    isHeadingWithoutExchangeSecType(contractSpec.secType) And _
    isHeadingWithoutCurrencySecType(contractSpec.secType) _
Then
    needHeadingRowWithSectypeOnly = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub selectContract( _
                ByVal Row As Long)
Const ProcName As String = "selectContract"
On Error GoTo Err

mSelectedRows.Add Row, CStr(Row)
highlightRow Row

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGrid()
Const ProcName As String = "setupGrid"
On Error GoTo Err

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

setupGridColumn 0, ContractsGridColumns.secType, ContractsGridColumnWidths.SecTypeWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupGridColumn 0, ContractsGridColumns.Exchange, ContractsGridColumnWidths.ExchangeWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter
setupGridColumn 0, ContractsGridColumns.CurrencyCode, ContractsGridColumnWidths.CurrencyWidth, True, TWControls40.AlignmentSettings.TwGridAlignCenterCenter
setupGridColumn 0, ContractsGridColumns.OptionRight, ContractsGridColumnWidths.OptionRightWidth, True, TWControls40.AlignmentSettings.TwGridAlignLeftCenter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGridColumn( _
                ByVal rowNumber As Long, _
                ByVal columnNumber As Long, _
                ByVal columnWidth As Single, _
                ByVal isLetters As Boolean, _
                ByVal align As TWControls40.AlignmentSettings)
Const ProcName As String = "setupGridColumn"
On Error GoTo Err

With TWGrid1
    .Row = rowNumber
    
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .colWidth(columnNumber) = 0
    End If
    
    .ColData(columnNumber) = columnNumber
    
    Dim lColumnWidth As Long
    If isLetters Then
        lColumnWidth = mLetterWidth * columnWidth
    Else
        lColumnWidth = mDigitWidth * columnWidth
    End If
    
    .colWidth(columnNumber) = lColumnWidth
    
    .ColAlignment(columnNumber) = align
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub toggleRowSelection( _
                ByVal Row As Long)
Const ProcName As String = "toggleRowSelection"
On Error GoTo Err

If isRowSelected(Row) Then
    deselectContract Row
Else
    selectContract Row
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function toneDown(ByVal pColor As Long) As Long
If (pColor And &H80000000) Then pColor = GetSysColor(pColor And &HFFFFFF)

toneDown = (((pColor And &HFF0000) / &H20000) And &HFF0000) + _
            (((pColor And &HFF00) / &H200) And &HFF00) + _
            ((pColor And &HFF) / &H2)
End Function

Private Sub writeHeadingRow( _
                ByVal contractSpec As IContractSpecifier)
Const ProcName As String = "writeHeadingRow"
On Error GoTo Err

Dim excludeExchange As Boolean
If isHeadingWithoutExchangeSecType(contractSpec.secType) Then excludeExchange = True

Dim excludeCurrency As Boolean
If isHeadingWithoutCurrencySecType(contractSpec.secType) Then excludeCurrency = True

TWGrid1.col = 0
TWGrid1.ColSel = ContractsGridColumns.MaxColumn
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
    TWGrid1.col = ContractsGridColumns.CurrencyCode
    TWGrid1.Text = contractSpec.CurrencyCode
End If

If Not excludeExchange Then
    TWGrid1.col = ContractsGridColumns.Exchange
    TWGrid1.Text = contractSpec.Exchange
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writeRow( _
                ByVal pContract As Contract)
Const ProcName As String = "writeRow"
On Error GoTo Err

TWGrid1.col = 0
TWGrid1.ColSel = ContractsGridColumns.MaxColumn
TWGrid1.CellBackColor = vbBlack ' Clear out any cell background colour
TWGrid1.CellFontBold = False
TWGrid1.CellForeColor = vbBlack

With pContract.Specifier
    TWGrid1.col = ContractsGridColumns.LocalSymbol
    TWGrid1.Text = .LocalSymbol
    
    If isFullHeadingSecType(.secType) Then
    Else
        If isHeadingWithoutExchangeSecType(.secType) Then
            TWGrid1.col = ContractsGridColumns.Exchange
            TWGrid1.Text = .Exchange
        End If
        If isHeadingWithoutCurrencySecType(.secType) Then
            TWGrid1.col = ContractsGridColumns.CurrencyCode
            TWGrid1.Text = .CurrencyCode
        End If
    End If
        
    'TWGrid1.col = ContractsGridColumns.Description
    'TWGrid1.Text = lContract.Description
    
    Select Case .secType
        Case SecTypeFuture
            TWGrid1.col = ContractsGridColumns.Expiry
            TWGrid1.Text = FormatDateTime(pContract.ExpiryDate, vbShortDate)
        Case SecTypeOption, SecTypeFuturesOption
            TWGrid1.col = ContractsGridColumns.Expiry
            TWGrid1.Text = FormatDateTime(pContract.ExpiryDate, vbShortDate)
        
            TWGrid1.col = ContractsGridColumns.OptionRight
            TWGrid1.Text = OptionRightToString(.Right)
            
            TWGrid1.col = ContractsGridColumns.Strike
            TWGrid1.Text = FormatPrice(.Strike, .secType, pContract.TickSize)
            TWGrid1.CellAlignment = TwGridAlignRightCenter
        'Case SecTypeCombo
    End Select
End With

mCount = mCount + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
