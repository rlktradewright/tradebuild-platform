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

Private mRegionRef                                  As WeakReference

Private mPriceFormatter                             As IPriceFormatter

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

Friend Property Let Hidden(ByVal Value As Boolean)
Const ProcName As String = "Hidden"

On Error GoTo Err

mHidden = Value

If mHidden Then removeLines

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Friend Property Let TextStyle( _
                ByVal Value As TextStyle)
Const ProcName As String = "TextStyle"

On Error GoTo Err

If Not mYGridTexts Is Nothing Then mYGridTexts.Style = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let YAxisRegion(ByVal Value As ChartRegion)
Const ProcName As String = "YAxisRegion"

On Error GoTo Err

If Value Is Nothing Then Exit Property

Set mYGridTexts = Value.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.Style = Value.YAxisTextStyle
mYGridTexts.Align = AlignBoxCentreLeft

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Function Generate( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Const ProcName As String = "Generate"

On Error GoTo Err

setHorizontalGridLines pBottom, pTop, pGridlineSpacing
setHorizontalGridTexts pBottom, pTop, pGridlineSpacing

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub Initialise( _
                ByVal pRegion As ChartRegion)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mRegionRef = CreateWeakReference(pRegion)

Set mYGridLines = regionObj.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
mYGridLines.Extended = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function FormatYValue(ByVal Y As Double) As String
Const ProcName As String = "FormatYValue"
On Error GoTo Err

FormatYValue = mPriceFormatter.FormatPrice(Y)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function

Private Sub removeLines()
Dim ln As Line
Const ProcName As String = "removeLines"

On Error GoTo Err

For Each ln In mYGridLines
    mYGridLines.RemoveObject ln
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHorizontalGridLines( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Dim Y As Double
Dim i As Long
Dim aLine As Line

Const ProcName As String = "setHorizontalGridLines"

On Error GoTo Err

If Not regionObj.HasGrid Then Exit Sub
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
For i = i + 1 To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition gNewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        gNewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setHorizontalGridTexts( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Dim Y As Double
Dim i As Long
Dim aText As Text

Const ProcName As String = "setHorizontalGridTexts"

On Error GoTo Err

If mYGridTexts Is Nothing Then Exit Sub

Y = Int(pBottom / pGridlineSpacing) * pGridlineSpacing
Do While Y < pTop
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    aText.Text = FormatYValue(Y)
    aText.Position = gNewPoint(15, Y, CoordsLogical, CoordsLogical)
    Y = Y + pGridlineSpacing
Loop
For i = i + 1 To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.Text = ""
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



