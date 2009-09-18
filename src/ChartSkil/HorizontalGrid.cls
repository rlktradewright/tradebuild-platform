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

Private mGridlineSpacingY                           As Double

Private mScaleGridSpacingY                          As Double
Private mYScaleFormatStr                            As String
Private mYScaleSubFormatStr                         As String

Private mHidden                                     As Boolean

Private mYScaleQuantum                              As Double

Private mIntegerYScale                              As Boolean

Private mRegionRef                                  As WeakReference
Private mYAxisRegionRef                             As WeakReference

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

Friend Property Let GridlineSpacingY(ByVal value As Double)
mGridlineSpacingY = value
End Property

Friend Property Let Hidden(ByVal value As Boolean)
mHidden = value
End Property

Friend Property Get Hidden() As Boolean
Hidden = mHidden
End Property

Friend Property Let IntegerYScale(ByVal value As Boolean)
mIntegerYScale = value
End Property

Friend Property Get ScaleGridSpacingY() As Single
ScaleGridSpacingY = mScaleGridSpacingY
End Property

Friend Property Let TextStyle( _
                ByVal value As TextStyle)
If Not mYGridTexts Is Nothing Then mYGridTexts.Style = value
End Property

Friend Property Let YAxisRegion(ByVal value As ChartRegion)
Set mYAxisRegionRef = CreateWeakReference(value)
Set mYGridTexts = value.AddGraphicObjectSeries(New TextSeries, LayerNumbers.LayerGridText)
mYGridTexts.Align = AlignBoxCentreLeft
End Property

Friend Property Let YScaleQuantum(ByVal value As Double)
mYScaleQuantum = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Generate( _
                ByVal pBottom As Double, _
                ByVal pTop As Double, _
                ByVal pHeightCm As Double)
calcGridSpacings pTop - pBottom, pHeightCm
setHorizontalGridLines pBottom, pTop
setHorizontalGridTexts pBottom, pTop
End Function

Friend Function GenerateYAxisCursorPositionText(ByVal Y As Single) As String
Dim fract As Single
Dim num32nds As Long

If mYScaleQuantum = 0.03125 Then
    fract = Y - Int(Y)
    num32nds = Int(fract / 0.03125)
    If mScaleGridSpacingY > 4 * 0.03125 Then
        GenerateYAxisCursorPositionText = Int(Y) & " ' " & Format(num32nds, "00")
    Else
        GenerateYAxisCursorPositionText = Int(Y) & " ' " & Format(num32nds, "00") & "." & Format(100 * (fract - num32nds * 0.03125) / 0.03125, "00")
    End If
Else
    GenerateYAxisCursorPositionText = Format(Y, mYScaleSubFormatStr)
End If
End Function

Friend Sub Initialise(ByVal Region As ChartRegion, pGridLineStyle As LineStyle)
Set mRegionRef = CreateWeakReference(Region)
Set mYGridLines = Region.AddGraphicObjectSeries(New LineSeries, LayerNumbers.LayerGrid)
mYGridLines.Style = pGridLineStyle
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcGridSpacings( _
                ByVal pHeight As Double, _
                ByVal pHeightCm As Double)
Dim numGridDivisions As Double
Dim gridSpacing As Double
Dim log10 As Double
Dim logSpacing As Double
Dim exp As Long
Dim mantissa As Double

numGridDivisions = pHeightCm / mGridlineSpacingY
gridSpacing = pHeight / numGridDivisions

If Not mHidden Then
    log10 = Log(10#)
    logSpacing = Log(gridSpacing) / log10
    
    exp = Fix(logSpacing)
    mantissa = Abs(logSpacing) - Abs(exp)
    
    If logSpacing < 0 Then
        exp = exp - 1
        mantissa = 1 - mantissa
    End If
    
    If mYScaleQuantum = 0.03125 And _
        exp < 0 _
    Then
        ' Tick size is 1/32, which influences where we put the horizontal grid lines
        If exp >= -1 Then
            If mantissa >= Log(5) / log10 Then
                mScaleGridSpacingY = 16 * 0.03125
            ElseIf mantissa >= Log(2.5) / log10 Then
                mScaleGridSpacingY = 8 * 0.03125
            ElseIf mantissa >= Log(1.25) / log10 Then
                mScaleGridSpacingY = 4 * 0.03125
            ElseIf mantissa >= Log(0.625) / log10 Then
                mScaleGridSpacingY = 2 * 0.03125
            Else
                mScaleGridSpacingY = 0.03125
            End If
        Else
            mScaleGridSpacingY = 0.03125
        End If
    Else
        If mantissa >= Log(5) / log10 Then
            mScaleGridSpacingY = 5 * 10 ^ exp
        ElseIf mantissa >= Log(2.5) / log10 Then
            mScaleGridSpacingY = 2.5 * 10 ^ exp
        ElseIf mantissa >= Log(2#) / log10 Then
            mScaleGridSpacingY = 2# * 10 ^ exp
        Else
            mScaleGridSpacingY = 10 ^ exp
        End If
    End If
    If mIntegerYScale Then
            mYScaleFormatStr = "0"
            mYScaleSubFormatStr = "0"
    Else
        If mScaleGridSpacingY < 1 Then
            mYScaleFormatStr = "0." & String(Abs(exp) + 1, "0")
            mYScaleSubFormatStr = "0." & String(Abs(exp) + 2, "0")
        Else
            mYScaleFormatStr = "0.0"
            mYScaleSubFormatStr = "0.00"
        End If
    End If
End If

End Sub

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function

Private Sub setHorizontalGridLines( _
                ByVal pBottom As Double, _
                ByVal pTop As Double)
Dim Y As Double
Dim i As Long
Dim aLine As ChartSkil26.Line
Dim reg As ChartRegion

If mHidden Then Exit Sub

Set reg = regionObj

Y = Int(pBottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While Y < pTop
    i = i + 1
    If i > mYGridLines.Count Then
        Set aLine = mYGridLines.Add
    Else
        Set aLine = mYGridLines.Item(i)
    End If
    aLine.SetPosition reg.NewPoint(0, Y, CoordsRelative, CoordsLogical), _
                        reg.NewPoint(100, Y, CoordsRelative, CoordsLogical)
    Y = Y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridLines.Count
    Set aLine = mYGridLines.Item(i)
    aLine.SetPosition reg.NewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        reg.NewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next
End Sub

Private Sub setHorizontalGridTexts( _
                ByVal pBottom As Double, _
                ByVal pTop As Double)
Dim Y As Double
Dim i As Long
Dim aText As Text
Dim reg As ChartRegion

If mYGridTexts Is Nothing Then Exit Sub

Set reg = yAxisRegionObj

Y = Int(pBottom / mScaleGridSpacingY) * mScaleGridSpacingY
Do While Y < pTop
    i = i + 1
    If i > mYGridTexts.Count Then
        Set aText = mYGridTexts.Add
    Else
        Set aText = mYGridTexts.Item(i)
    End If
    If mYScaleQuantum = 0.03125 Then
        ' Tick size is 1/32 - label accordingly
        aText.Text = Int(Y) & " ' " & Format((Y - Int(Y)) / 0.03125, "00")
    Else
        aText.Text = Format(Y, mYScaleFormatStr)
    End If
    aText.position = reg.NewPoint(15, Y, CoordsLogical, CoordsLogical)
    Y = Y + mScaleGridSpacingY
Loop
For i = i + 1 To mYGridTexts.Count
    Set aText = mYGridTexts.Item(i)
    aText.Text = ""
Next
End Sub

Private Function yAxisRegionObj() As ChartRegion
Set yAxisRegionObj = mYAxisRegionRef.Target
End Function



