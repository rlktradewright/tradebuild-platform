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

Private mXGridLabelsCoarse                          As TextSeries
Private mXGridLabels                                As TextSeries
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
    mXGridLabels.Align = AlignBoxBottomCentre
    mXGridLabelsCoarse.Align = AlignBoxBottomCentre
Case XGridTextPositionBottom
    mXGridLabels.Align = AlignBoxBottomCentre
    mXGridLabelsCoarse.Align = AlignBoxBottomCentre
Case XGridTextPositionTop
    mXGridLabels.Align = AlignBoxTopCentre
    mXGridLabelsCoarse.Align = AlignBoxTopCentre
End Select
End Property

Friend Property Let GridTextStyle( _
                ByVal Value As TextStyle)
Const ProcName As String = "GridTextStyle"
On Error GoTo Err

If Not mXGridLabels Is Nothing Then mXGridLabels.Style = Value
If Not mXGridLabelsCoarse Is Nothing Then mXGridLabelsCoarse.Style = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGrid(ByVal Value As Boolean)
Const ProcName As String = "HasGrid"
On Error GoTo Err

mHasGrid = Value
If Not mHasGrid Then clearRemainingGridLines mXGridLines.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let HasGridText(ByVal Value As Boolean)
Const ProcName As String = "HasGridText"
On Error GoTo Err

mHasGridText = Value
If Not mHasGridText Then
    clearRemainingGridLabels mXGridLabels.Enumerator
    clearRemainingCoarseGridLabels mXGridLabelsCoarse.Enumerator
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

If Not mHasGrid And Not mHasGridText Then Exit Sub

Static lPeriodsNeedingGridlines As New EnumerableCollection
Static lPeriodsNeedingGridTexts As New EnumerableCollection
Static lPeriodsNeedingCoarseGridTexts As New EnumerableCollection

Dim startPeriod As Long
startPeriod = Int(IIf(mViewport.Boundary.Left < 1, _
                    1, _
                    Int(mViewport.Boundary.Left)))

Dim endPeriod As Long
endPeriod = Int(IIf(mViewport.Boundary.Right > mPeriods.Count, _
                mPeriods.Count, _
                Int(mViewport.Boundary.Right)))
If endPeriod < startPeriod Then Exit Sub

Dim i As Long
For i = startPeriod To endPeriod
    Dim lPeriod As Period
    Set lPeriod = mPeriods(i)
    
    If lPeriod.VerticalGridLineType = VerticalGridlineTypeNone Then
    ElseIf mHasGrid Then
        lPeriodsNeedingGridlines.Add lPeriod
    End If
    
    If mHasGridText Then
        If lPeriod.Labels.CoarseLabel <> "" Then lPeriodsNeedingCoarseGridTexts.Add lPeriod
        If lPeriod.Labels.Label <> "" Then lPeriodsNeedingGridTexts.Add lPeriod
    End If
Next

generateGridLines lPeriodsNeedingGridlines.Enumerator
lPeriodsNeedingGridlines.Clear

generateGridLabels lPeriodsNeedingGridTexts.Enumerator
lPeriodsNeedingGridTexts.Clear
generateCoarseGridLabels lPeriodsNeedingCoarseGridTexts.Enumerator
lPeriodsNeedingCoarseGridTexts.Clear

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
mXGridLines.Layer = LayerNumbers.LayerGrid

Set mXGridLabels = pTexts
mXGridLabels.Layer = LayerNumbers.LayerGridText + 1
mXGridLabels.Align = AlignBoxBottomCentre
mXGridLabels.FixedX = False
mXGridLabels.FixedY = True
mXGridLabels.HideIfBlank = True
mXGridLabels.MultiLine = True
mXGridLabels.Justification = JustifyCentre

Set mXGridLabelsCoarse = pCoarseTexts
mXGridLabelsCoarse.Layer = LayerNumbers.LayerGridText
mXGridLabelsCoarse.Align = AlignBoxTopCentre
mXGridLabelsCoarse.FixedX = False
mXGridLabelsCoarse.FixedY = True
mXGridLabelsCoarse.HideIfBlank = True
mXGridLabelsCoarse.MultiLine = True
mXGridLabelsCoarse.Justification = JustifyCentre

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearRemainingGridLines(ByVal pEn As Enumerator)
Const ProcName As String = "clearRemainingGridLines"
On Error GoTo Err

Do While pEn.MoveNext
    Dim aLine As ChartSkil27.Line
    Set aLine = pEn.Current
    aLine.SetPosition gNewPoint(-1000, 0, CoordsLogical, CoordsLogical), _
                        gNewPoint(-1000, 0, CoordsLogical, CoordsLogical)
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearRemainingCoarseGridLabels(ByVal pEn As Enumerator)
Const ProcName As String = "clearRemainingCoarseGridLabels"
On Error GoTo Err

Do While pEn.MoveNext
    Dim aText As Text
    Set aText = pEn.Current
    aText.Text = ""
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearRemainingGridLabels(ByVal pEn As Enumerator)
Const ProcName As String = "clearRemainingGridLabels"
On Error GoTo Err

Do While pEn.MoveNext
    Dim aText As Text
    Set aText = pEn.Current
    aText.Text = ""
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateGridLabels(ByVal pPeriodsEn As Enumerator)
Const ProcName As String = "generateGridLabels"
On Error GoTo Err

Dim lPrevLabel As Text

Dim lTextsEn As Enumerator
Set lTextsEn = mXGridLabels.Enumerator

Dim lPeriod As Period
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    If Not lTextsEn.MoveNext Then
        ' we don't currently have enough grid labels
        Exit Do
    End If
    setupGridLabel lPeriod, lTextsEn.Current, lPrevLabel
    Set lPeriod = Nothing
Loop

If lPeriod Is Nothing Then
    clearRemainingGridLabels lTextsEn
    Exit Sub
End If

' there are still some periods that don't have a grid label
setupGridLabel lPeriod, mXGridLabels.Add, lPrevLabel
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    setupGridLabel lPeriod, mXGridLabels.Add, lPrevLabel
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateCoarseGridLabels(ByVal pPeriodsEn As Enumerator)
Const ProcName As String = "generateCoarseGridLabels"
On Error GoTo Err

Dim lTextsEn As Enumerator
Set lTextsEn = mXGridLabelsCoarse.Enumerator

Dim lPeriod As Period
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    If Not lTextsEn.MoveNext Then
        ' we don't currently have enough coarse grid labels
        Exit Do
    End If
    setupGridLabelCoarse lPeriod, lTextsEn.Current
    Set lPeriod = Nothing
Loop

If lPeriod Is Nothing Then
    clearRemainingCoarseGridLabels lTextsEn
    Exit Sub
End If

' there are still some periods that don't have a coarse grid label
setupGridLabelCoarse lPeriod, mXGridLabels.Add
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    setupGridLabelCoarse lPeriod, mXGridLabels.Add
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub generateGridLines(ByVal pPeriodsEn As Enumerator)
Const ProcName As String = "generateGridLines"
On Error GoTo Err

Dim lLinesEn As Enumerator
Set lLinesEn = mXGridLines.Enumerator

Dim lPeriod As Period
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    If Not lLinesEn.MoveNext Then
        ' we don't currently have enough grid lines
        Exit Do
    End If
    setupGridLine lPeriod, lLinesEn.Current
    Set lPeriod = Nothing
Loop

If lPeriod Is Nothing Then
    clearRemainingGridLines lLinesEn
    Exit Sub
End If

' there are still some periods that don't have a grid line
setupGridLine lPeriod, mXGridLines.Add
Do While pPeriodsEn.MoveNext
    Set lPeriod = pPeriodsEn.Current
    setupGridLine lPeriod, mXGridLines.Add
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Sub setupGridLabel(ByVal pPeriod As Period, ByVal pText As Text, ByRef pPrevLabel As Text)
Const ProcName As String = "setupGridLabel"
On Error GoTo Err

Select Case mGridTextPosition
Case XGridTextPositionCentre
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                50, _
                                CoordsLogical, _
                                CoordsRelative)
Case XGridTextPositionBottom
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsDistance)
Case XGridTextPositionTop
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsCounterDistance)
End Select

pText.Text = pPeriod.Labels.Label

If gGraphicObjectsOverlap(pText, pPrevLabel) Then
    pText.Text = ""
Else
    Set pPrevLabel = pText
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGridLabelCoarse(ByVal pPeriod As Period, ByVal pText As Text)
Const ProcName As String = "setupGridLabelCoarse"
On Error GoTo Err

Select Case mGridTextPosition
Case XGridTextPositionCentre
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                50, _
                                CoordsLogical, _
                                CoordsRelative)
    pText.Text = pPeriod.Labels.CoarseLabel & vbCrLf
Case XGridTextPositionBottom
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsDistance)
    pText.Text = pPeriod.Labels.CoarseLabel & vbCrLf
Case XGridTextPositionTop
    pText.Position = gNewPoint(pPeriod.PeriodNumber - 0.5, _
                                0#, _
                                CoordsLogical, _
                                CoordsCounterDistance)
    pText.Text = vbCrLf & pPeriod.Labels.CoarseLabel
End Select
pText.Text = pPeriod.Labels.CoarseLabel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupGridLine(ByVal pPeriod As Period, ByVal pLine As ChartSkil27.Line)
Const ProcName As String = "setupGridLine"
On Error GoTo Err

pLine.LocalStyle = getLineStyle(pPeriod.VerticalGridLineType)

pLine.SetPosition gNewPoint(pPeriod.PeriodNumber - 0.5, _
                            0.1, _
                            CoordsLogical, _
                            CoordsDistance), _
                    gNewPoint(pPeriod.PeriodNumber - 0.5, _
                            0.1, _
                            CoordsLogical, _
                            CoordsCounterDistance)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

