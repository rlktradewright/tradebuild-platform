VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GridManager"

Private Const CellBackColorOdd As Long = &HF8F8F8
Private Const CellBackColorEven As Long = &HEEEEEE

Private Const TestDataGridRowsInitial As Long = 2
Private Const TestDataGridRowsIncrement As Long = 25

'@================================================================================
' Member variables
'@================================================================================

Private mGrid As MSFlexGrid
Private mParent As Form
Private mPriceFormatString As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Cols(ByVal Value As Long)
mGrid.Cols = Value
End Property

Friend Property Let PriceFormatString(ByVal Value As String)
mPriceFormatString = Value
End Property

Friend Property Let Redraw(ByVal Value As Boolean)
mGrid.Redraw = Value
End Property

Friend Property Let Row(ByVal Value As Long)
If Value > mGrid.Rows - 1 Then mGrid.Rows = mGrid.Rows + TestDataGridRowsIncrement
mGrid.Row = Value
End Property

Friend Property Get Row() As Long
Row = mGrid.Row
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pGrid As MSFlexGrid)
Set mGrid = pGrid
Set mParent = mGrid.Parent
End Sub

Friend Sub SetCell(ByVal pColIndex As Long, ByVal pValue As String)
mGrid.Col = pColIndex
mGrid.Text = pValue
adjustGridColWidth pColIndex, calcTextWidth(pValue)
setCellBackColor pColIndex
End Sub

Friend Sub SetCellDate(ByVal pColIndex As Long, ByVal pValue As Date)
SetCell pColIndex, FormatTimestamp(pValue, TimestampDateAndTimeISO8601 + TimestampNoMillisecs)
End Sub

Friend Sub SetCellDouble(ByVal pColIndex As Long, ByVal pValue As Double)
SetCell pColIndex, CStr(pValue)
End Sub

Friend Sub SetCellLong(ByVal pColIndex As Long, ByVal pValue As Long)
SetCell pColIndex, CStr(pValue)
End Sub

Friend Sub SetCellPrice(ByVal pColIndex As Long, ByVal pValue As Double)
SetCell pColIndex, Format(pValue, mPriceFormatString)
End Sub

Friend Sub SetupDataColumns()

Const ProcName As String = "SetupDataColumns"
On Error GoTo Err

With mGrid
    .Clear
    .AllowBigSelection = True
    .AllowUserResizing = flexResizeBoth
    .FillStyle = flexFillRepeat
    .FocusRect = flexFocusNone
    .HighLight = flexHighlightNever
    
    .Cols = TestDataGridColumns.StudyValue1
    .Rows = TestDataGridRowsInitial
    .FixedRows = 1
    .FixedCols = 0
End With
    
setupGridColumn TestDataGridColumns.BarNumber, "#", AlignmentSettings.flexAlignLeftCenter
setupGridColumn TestDataGridColumns.TimeStamp, "Timestamp", AlignmentSettings.flexAlignLeftCenter
setupGridColumn TestDataGridColumns.OpenValue, "Open", AlignmentSettings.flexAlignRightCenter
setupGridColumn TestDataGridColumns.HighValue, "High", AlignmentSettings.flexAlignRightCenter
setupGridColumn TestDataGridColumns.LowValue, "Low", AlignmentSettings.flexAlignRightCenter
setupGridColumn TestDataGridColumns.CloseValue, "Close", AlignmentSettings.flexAlignRightCenter
setupGridColumn TestDataGridColumns.Volume, "Volume", AlignmentSettings.flexAlignRightCenter
setupGridColumn TestDataGridColumns.OpenInterest, "Open Interest", AlignmentSettings.flexAlignRightCenter

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetupStudyValueColumns( _
                ByVal studyConfig As StudyConfiguration)
Dim svd As StudyValueDefinition
Dim svc As StudyValueConfiguration
Dim i As Long
Dim j As Long

' remove any existing study Value columns
Const ProcName As String = "SetupStudyValueColumns"
On Error GoTo Err

mGrid.Cols = TestDataGridColumns.StudyValue1

For i = 1 To studyConfig.StudyValueConfigurations.Count
    Set svc = studyConfig.StudyValueConfigurations.Item(i)
    If svc.IncludeInChart Then
        Set svd = studyConfig.Study.StudyDefinition.StudyValueDefinitions.Item(svc.ValueName)
        setupGridColumn TestDataGridColumns.StudyValue1 + j, _
                                svd.Name, _
                                getAlignment(svd)
        j = j + 1
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getAlignment(ByVal pSvd As StudyValueDefinition) As AlignmentSettings
Select Case pSvd.ValueMode
Case ValueModeNone
    Select Case pSvd.ValueType
    Case ValueTypeNone
        getAlignment = flexAlignLeftCenter
    Case ValueTypeInteger
        getAlignment = flexAlignRightCenter
    Case ValueTypeReal
        getAlignment = flexAlignRightCenter
    Case ValueTypeString
        getAlignment = flexAlignLeftCenter
    Case ValueTypeDate
        getAlignment = flexAlignCenterCenter
    Case ValueTypeBoolean
        getAlignment = flexAlignCenterCenter
    End Select
Case ValueModeLine
    getAlignment = flexAlignLeftCenter
Case ValueModeBar
    getAlignment = flexAlignLeftCenter
Case ValueModeText
    getAlignment = flexAlignLeftCenter
End Select
End Function

Private Sub adjustGridColWidth(ByVal pColIndex As Long, ByVal pMinWidth As Long)
If pMinWidth > mGrid.ColWidth(pColIndex) Then mGrid.ColWidth(pColIndex) = pMinWidth
End Sub

Private Function calcTextWidth(ByRef pText As String) As Long
Const ProcName As String = "calcTextWidth"
On Error GoTo Err

calcTextWidth = mParent.TextWidth(pText) + 6 * Screen.TwipsPerPixelX
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setCellBackColor( _
                ByVal pcolumn As Long)

Const ProcName As String = "setCellBackColor"
On Error GoTo Err

mGrid.Col = pcolumn
mGrid.CellBackColor = IIf(mGrid.Row Mod 2 = 0, CellBackColorEven, CellBackColorOdd)
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGridColumn( _
                ByVal columnNumber As Long, _
                ByVal columnHeader As String, _
                ByVal align As AlignmentSettings)
Const ProcName As String = "setupGridColumn"
On Error GoTo Err

With mGrid
    If (columnNumber + 1) > .Cols Then
        .Cols = columnNumber + 1
        .ColWidth(columnNumber) = 0
    End If
    
    .ColWidth(columnNumber) = calcTextWidth(columnHeader)
        
    .ColAlignment(columnNumber) = align
    .FixedAlignment(columnNumber) = AlignmentSettings.flexAlignCenterCenter
    .Row = 0
    .Col = columnNumber
    .Text = columnHeader
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                


