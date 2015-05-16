VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VerticalGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "VerticalGrid"

'@================================================================================
' Member variables
'@================================================================================

Private mRegionRef                                  As WeakReference

Private mXGridTextsCoarse                           As TextSeries
Private mXGridTexts                                 As TextSeries
Private mXGridLines                                 As LineSeries

Private mSessionEndGridLineStyle                    As LineStyle
Private mSessionStartGridLineStyle                  As LineStyle

Private mPeriods                                    As Periods
Private mViewport                                   As ViewPort

Private mHasGrid                                    As Boolean
Private mHasGridText                                As Boolean

Private mGridTextPosition                           As XGridTextPositions

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
mXGridLines.Style = Value
End Property

Friend Property Let GridTextPosition(ByVal Value As XGridTextPositions)
mGridTextPosition = Value
Select Case mGridTextPosition
Case XGridTextPositionCentre
    mXGridTexts.Align = AlignBoxBottomCentre
    mXGridTextsCoarse.Align = AlignBoxBottomCentre
Case XGridTextPositionBottom
    mXGridTexts.Align = AlignBoxBottomCentre
    mXGridTextsCoarse.Align = AlignBoxBottomCentre
Case XGridTextPositionTop
    mXGridTexts.Align = AlignBoxTopCentre
    mXGridTextsCoarse.Align = AlignBoxTopCentre
End Select
End Property

Friend Property Let GridTextStyle( _
                ByVal Value As TextStyle)
Const ProcName As String = "GridTextStyle"
On Error GoTo Err

If Not mXGridTexts Is Nothing Then mXGridTexts.Style = Value
If Not mXGridTextsCoarse Is Nothing Then mXGridTextsCoarse.Style = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGrid(ByVal Value As Boolean)
Const ProcName As String = "HasGrid"
On Error GoTo Err

mHasGrid = Value
If Not mHasGrid Then clearGridLinesFrom 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGridText(ByVal Value As Boolean)
Const ProcName As String = "HasGridText"
On Error GoTo Err

mHasGridText = Value
If Not mHasGridText Then
    clearGridLabelsFrom 1
    clearCoarseGridLabelsFrom 1
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let SessionEndGridLineStyle(ByVal Value As LineStyle)
Set mSessionEndGridLineStyle = Value
End Property

Friend Property Let SessionStartGridLineStyle(ByVal Value As LineStyle)
Set mSessionStartGridLineStyle = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Generate()
Const ProcName As String = "Generate"
On Error GoTo Err

Static sPrevLabel As Text
Dim lPeriod As Period
Dim i As Long
Dim j As Long
Dim k As Long
Dim l As Long
Dim lLabel As Text

If Not mHasGrid And Not mHasGridText Then Exit Sub

Set sPrevLabel = Nothing

For i = Int(IIf(mViewport.Boundary.Left < 1, _
                1, _
                Int(mViewport.Boundary.Left))) _
    To Int(IIf(mViewport.Boundary.Right > mPeriods.Count, _
                mPeriods.Count, _
                Int(mViewport.Boundary.Right)))

    Set lPeriod = mPeriods(i)
    
    If lPeriod.VerticalGridLineType = VerticalGridlineTypeNone Then
        If mHasGridText And lPeriod.Labels.CoarseLabel <> "" Then
            l = l + 1
            setupGridLabelCoarse lPeriod, l
        End If
    Else
        If mHasGrid Then
            j = j + 1
            setupGridLine lPeriod, j
        End If
        
        If mHasGridText Then
            If lPeriod.Labels.CoarseLabel <> "" Then
                l = l + 1
                setupGridLabelCoarse lPeriod, l
            End If
            
            If lPeriod.Labels.Label <> "" Then
                k = k + 1
                Set lLabel = setupGridLabel(lPeriod, k)
                If gGraphicObjectsOverlap(lLabel, sPrevLabel) Then
                    lLabel.Text = ""
                Else
                    Set sPrevLabel = lLabel
                End If
            End If
        End If
    End If
Next

clearGridLinesFrom j + 1
clearGridLabelsFrom k + 1
clearCoarseGridLabelsFrom l + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pLines As LineSeries, _
                ByVal pTexts As TextSeries, _
                ByVal pCoarseTexts As TextSeries, _
                ByVal pPeriods As Periods, _
                ByVal pViewport As ViewPort)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mPeriods = pPeriods
Set mViewport = pViewport

Set mXGridLines = pLines
mXGridLines.layer = LayerNumbers.LayerGrid

Set mXGridTexts = pTexts
mXGridTexts.layer = LayerNumbers.LayerGridText + 1
mXGridTexts.Align = AlignBoxBottomCentre
mXGridTexts.FixedX = False
mXGridTexts.FixedY = True
mXGridTexts.HideIfBlank = True
mXGridTexts.MultiLine = True
mXGridTexts.Justification = JustifyCentre

Set mXGridTextsCoarse = pCoarseTexts
mXGridTextsCoarse.layer = LayerNumbers.LayerGridText
mXGridTextsCoarse.Align = AlignBoxTopCentre
mXGridTextsCoarse.FixedX = False
mXGridTextsCoarse.FixedY = True
mXGridTextsCoarse.HideIfBlank = True
mXGridTextsCoarse.MultiLine = True
mXGridTextsCoarse.Justification = JustifyCentre

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearGridLinesFrom(ByVal pIndex As Long)
Const ProcName As String = "clearGridLinesFrom"
On Error GoTo Err

Dim i As Long
Dim aLine As ChartSkil27.Line

For i = pIndex To mXGridLines.Count
    Set aLine = mXGridLines.Item(i)
    aLine.SetPosition gNewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        gNewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearCoarseGridLabelsFrom(ByVal pIndex As Long)
Const ProcName As String = "clearCoarseGridLabelsFrom"
On Error GoTo Err

Dim i As Long
Dim aText As Text

For i = pIndex To mXGridTextsCoarse.Count
    Set aText = mXGridTextsCoarse.Item(i)
    aText.Text = ""
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearGridLabelsFrom(ByVal pIndex As Long)
Const ProcName As String = "clearGridLabelsFrom"
On Error GoTo Err

Dim i As Long
Dim aText As Text

For i = pIndex To mXGridTexts.Count
    Set aText = mXGridTexts.Item(i)
    aText.Text = ""
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function GetGraphicObject(ByVal pSeries As IGraphicObjectSeries, ByVal pIndex As Long) As IGraphicObject
Const ProcName As String = "getGraphicObject"
On Error GoTo Err

If pIndex <= pSeries.Count Then
    Set GetGraphicObject = pSeries.Item(pIndex)
Else
    Set GetGraphicObject = pSeries.Add
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getLineStyle(ByVal pGridLineType As VerticalGridlineTypes) As LineStyle
Select Case pGridLineType
Case VerticalGridlineTypePeriodBoundary
    ' use default style
Case VerticalGridlineTypeSessionEnd
    Set getLineStyle = mSessionEndGridLineStyle
Case VerticalGridlineTypeSessionStart
    Set getLineStyle = mSessionStartGridLineStyle
Case VerticalGridlineTypeDateChange
    ' use default style
End Select
End Function

Private Sub removeUnusedGraphicObjects(ByVal pSeries As IGraphicObjectSeries, ByVal pIndex As Long)
Do While pSeries.Count >= pIndex
    pSeries.Remove pSeries.Count
Loop
End Sub

Private Function setupGridLabel(ByVal pPeriod As Period, ByVal pIndex As Long) As Text
Dim lText As Text

Set lText = GetGraphicObject(mXGridTexts, pIndex)
lText.Text = pPeriod.Labels.Label

Select Case mGridTextPosition
Case XGridTextPositionCentre
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                50, _
                                CoordsLogical, _
                                CoordsRelative)
Case XGridTextPositionBottom
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsDistance)
Case XGridTextPositionTop
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsCounterDistance)
End Select

Set setupGridLabel = lText

End Function

Private Sub setupGridLabelCoarse(ByVal pPeriod As Period, ByVal pIndex As Long)
Dim lText As Text

Set lText = GetGraphicObject(mXGridTextsCoarse, pIndex)

Select Case mGridTextPosition
Case XGridTextPositionCentre
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                50, _
                                CoordsLogical, _
                                CoordsRelative)
    lText.Text = pPeriod.Labels.CoarseLabel & vbCrLf
Case XGridTextPositionBottom
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsDistance)
    lText.Text = pPeriod.Labels.CoarseLabel & vbCrLf
Case XGridTextPositionTop
    lText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsCounterDistance)
    lText.Text = vbCrLf & pPeriod.Labels.CoarseLabel
End Select
End Sub

Private Sub setupGridLine(ByVal pPeriod As Period, ByVal pIndex As Long)
Dim lLine As ChartSkil27.Line

Set lLine = GetGraphicObject(mXGridLines, pIndex)

lLine.LocalStyle = getLineStyle(pPeriod.VerticalGridLineType)

lLine.SetPosition gNewPoint(pPeriod.PeriodNumber - 0.5, _
                            0.1, _
                            CoordsLogical, _
                            CoordsDistance), _
                    gNewPoint(pPeriod.PeriodNumber - 0.5, _
                            0.1, _
                            CoordsLogical, _
                            CoordsCounterDistance)
End Sub

