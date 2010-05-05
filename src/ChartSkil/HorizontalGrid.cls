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
Private mYAxisRegionRef                             As WeakReference

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

Friend Property Let GridLineStyle(ByVal value As LineStyle)
mYGridLines.Style = value
End Property

Friend Property Let Hidden(ByVal value As Boolean)
Const ProcName As String = "Hidden"
Dim failpoint As String
On Error GoTo Err

mHidden = value

If mHidden Then removeLines

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Friend Property Let PriceFormatter(ByVal value As IPriceFormatter)
Set mPriceFormatter = value
End Property

Friend Property Get PriceFormatter() As IPriceFormatter
Set PriceFormatter = mPriceFormatter
End Property

Friend Property Let TextStyle( _
                ByVal value As TextStyle)
Const ProcName As String = "TextStyle"
Dim failpoint As String
On Error GoTo Err

If Not mYGridTexts Is Nothing Then mYGridTexts.Style = value

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)
Const ProcName As String = "YAxisRegion"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing Then Exit Property

Set mYAxisRegionRef = CreateWeakReference(value)
Set mYGridTexts = yAxisRegionObj.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.Align = AlignBoxCentreLeft

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
Dim failpoint As String
On Error GoTo Err

removeLines
Set mYGridTexts = Nothing
Set mYGridLines = Nothing

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Function Generate( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Const ProcName As String = "Generate"
Dim failpoint As String
On Error GoTo Err

setHorizontalGridLines pBottom, pTop, pGridlineSpacing
setHorizontalGridTexts pBottom, pTop, pGridlineSpacing

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal pRegion As ChartRegion)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mRegionRef = CreateWeakReference(pRegion)

Set mYGridLines = regionObj.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
mYGridLines.Extended = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function

Private Sub removeLines()
Dim ln As ChartSkil26.Line
Const ProcName As String = "removeLines"
Dim failpoint As String
On Error GoTo Err

For Each ln In mYGridLines
    mYGridLines.RemoveObject ln
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setHorizontalGridLines( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Dim Y As Double
Dim i As Long
Dim aLine As ChartSkil26.Line
Dim reg As ChartRegion

Const ProcName As String = "setHorizontalGridLines"
Dim failpoint As String
On Error GoTo Err

If mHidden Then Exit Sub

Set reg = regionObj

Y = Int(pBottom / pGridlineSpacing) * pGridlineSpacing
Do While Y < pTop
    i = i + 1
    If i > mYGridLines.Count Then
        Set aLine = mYGridLines.Add
    Else
        Set aLine = mYGridLines.Item(i)
    End If
    aLine.SetPosition reg.NewPoint(0, Y, CoordsRelative, CoordsLogical), _
                        reg.NewPoint(100, Y, CoordsRelative, CoordsLogical)
    Y = Y + pGridlineSpacing
Loop
For i = i + 1 To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition reg.NewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        reg.NewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Sub setHorizontalGridTexts( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pGridlineSpacing As Double)
Dim Y As Double
Dim i As Long
Dim aText As Text
Dim reg As ChartRegion

Const ProcName As String = "setHorizontalGridTexts"
Dim failpoint As String
On Error GoTo Err

If mYGridTexts Is Nothing Then Exit Sub

Set reg = yAxisRegionObj

Y = Int(pBottom / pGridlineSpacing) * pGridlineSpacing
Do While Y < pTop
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    aText.Text = FormatYValue(Y)
    aText.Position = reg.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    Y = Y + pGridlineSpacing
Loop
For i = i + 1 To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.Text = ""
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Private Function yAxisRegionObj() As ChartRegion
Set yAxisRegionObj = mYAxisRegionRef.Target
End Function


