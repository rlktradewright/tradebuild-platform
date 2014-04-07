VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HorizontalGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "HorizontalGrid"

'@================================================================================
' Member variables
'@================================================================================

Private mYGridTexts                                 As TextSeries
Private mYGridLines                                 As LineSeries

Private mHidden                                     As Boolean

Private mPriceFormatter                             As IPriceFormatter
Private mViewport                                   As ViewPort

Private mHasGrid                                    As Boolean
Private mHasGridText                                As Boolean

Private mGridTextPosition                           As YGridTextPositions

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

Friend Property Let GridLineStyle(ByVal Value As LineStyle)
mYGridLines.Style = Value
End Property

Friend Property Let GridTextPosition(ByVal Value As YGridTextPositions)
mGridTextPosition = Value
Select Case mGridTextPosition
Case YGridTextPositionCentre
    mYGridTexts.Align = AlignBoxCentreCentre
Case YGridTextPositionLeft
    mYGridTexts.Align = AlignBoxCentreLeft
Case YGridTextPositionRight
    mYGridTexts.Align = AlignCentreRight
End Select
End Property

Friend Property Let GridTextStyle( _
                ByVal Value As TextStyle)
Const ProcName As String = "TextStyle"

On Error GoTo Err

If Not mYGridTexts Is Nothing Then mYGridTexts.Style = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGrid(ByVal Value As Boolean)
Const ProcName As String = "HasGrid"
On Error GoTo Err

mHasGrid = Value
If Not mHasGrid Then clearLinesFrom 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGridText(ByVal Value As Boolean)
Const ProcName As String = "HasGridText"
On Error GoTo Err

mHasGridText = Value
If Not mHasGridText Then clearTextsFrom 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Hidden(ByVal Value As Boolean)
Const ProcName As String = "Hidden"

On Error GoTo Err

mHidden = Value

If mHidden Then
    clearLinesFrom 1
    clearTextsFrom 1
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Friend Property Let PriceFormatter(ByVal Value As IPriceFormatter)
Set mPriceFormatter = Value
End Property

Friend Property Get PriceFormatter() As IPriceFormatter
Set PriceFormatter = mPriceFormatter
End Property

Friend Property Let YAxisRegion(ByVal Value As ChartRegion)
Const ProcName As String = "YAxisRegion"

On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mYGridTexts = Value.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.Align = AlignBoxCentreLeft

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

removeLines
removeTexts
Set mYGridTexts = Nothing
Set mYGridLines = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Generate(ByVal pGridlineSpacing As Double)
Const ProcName As String = "Generate"

On Error GoTo Err

setHorizontalGridLines mViewport.Bottom, mViewport.Top, pGridlineSpacing
setHorizontalGridTexts mViewport.Bottom, mViewport.Top, pGridlineSpacing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pLines As LineSeries, _
                ByVal pTexts As TextSeries, _
                ByVal pViewport As ViewPort)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mYGridLines = pLines
mYGridLines.layer = LayerNumbers.LayerGrid
'mYGridLines.Extended = True

Set mYGridTexts = pTexts
mYGridTexts.layer = LayerNumbers.LayerGridText
mYGridTexts.Align = AlignBoxCentreLeft
mYGridTexts.HideIfBlank = True

Set mViewport = pViewport

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearLinesFrom(ByVal pIndex As Long)
Const ProcName As String = "clearLinesFrom"
On Error GoTo Err

Dim i As Long
Dim aLine As ChartSkil27.Line

For i = pIndex To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition gNewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        gNewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearTextsFrom(ByVal pIndex As Long)
Const ProcName As String = "clearTextsFrom"
On Error GoTo Err

Dim i As Long
Dim aText As Text

For i = pIndex To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.Text = ""
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function FormatYValue(ByVal Y As Double) As String
Const ProcName As String = "FormatYValue"
On Error GoTo Err

FormatYValue = mPriceFormatter.FormatPrice(Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeLines()
Dim ln As ChartSkil27.Line
Const ProcName As String = "removeLines"

On Error GoTo Err

For Each ln In mYGridLines
    mYGridLines.RemoveObject ln
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeTexts()
Dim tx As Text
Const ProcName As String = "removeTexts"

On Error GoTo Err

For Each tx In mYGridTexts
    mYGridTexts.RemoveObject tx
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHorizontalGridLines( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Const ProcName As String = "setHorizontalGridLines"
On Error GoTo Err

Dim Y As Double
Dim i As Long
Dim aLine As ChartSkil27.Line

If Not mHasGrid Then Exit Sub
If mHidden Then Exit Sub

Y = Int(pBottom / pGridlineSpacing) * pGridlineSpacing
Do While Y < pTop
    i = i + 1
    If i > mYGridLines.Count Then
        Set aLine = mYGridLines.Add
    Else
        Set aLine = mYGridLines.Item(i)
    End If
    aLine.SetPosition gNewPoint(0, Y, CoordsRelative, CoordsLogical), _
                        gNewPoint(100, Y, CoordsRelative, CoordsLogical)
    Y = Y + pGridlineSpacing
Loop

clearLinesFrom i + 1

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrOverflow Then
    gLogger.Log "Overflow: Y=" & CStr(Y) & "; pBottom=" & CStr(pBottom) & "; pTop=" & CStr(pTop), ProcName, ModuleName, LogLevelSevere
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setHorizontalGridTexts( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Const ProcName As String = "setHorizontalGridTexts"
On Error GoTo Err

Dim Y As Double
Dim i As Long
Dim aText As Text

If Not mHasGridText Then Exit Sub

Y = Int(pBottom / pGridlineSpacing) * pGridlineSpacing
Do While Y < pTop
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    aText.Text = FormatYValue(Y)
    
    Select Case mGridTextPosition
    Case YGridTextPositionCentre
        aText.Position = gNewPoint(50, Y, CoordsRelative, CoordsLogical)
    Case YGridTextPositionLeft
        aText.Position = gNewPoint(0.2, Y, CoordsDistance, CoordsLogical)
    Case YGridTextPositionRight
        aText.Position = gNewPoint(0.2, Y, CoordsCounterDistance, CoordsLogical)
    End Select
    
    Y = Y + pGridlineSpacing
Loop

clearTextsFrom i + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



