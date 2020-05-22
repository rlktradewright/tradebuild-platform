VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeferredRegionPainter"
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

Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum DeferredPaintActions
    DeferredPaintRepaintNone
    DeferredPaintPaintAreas
    DeferredPaintPaint
    DeferredPaintChangeGaugeAndPaint
    DeferredPaintScaleAndPaint
    DeferredPaintCalculateAndPaint
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type RedrawObjectsTableEntry
    graphObj            As IGraphicObject
    areaToRedraw        As TRectangle
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DeferredRegionPainter"

'@================================================================================
' Member variables
'@================================================================================

Private mRegion As ChartRegion

Private mDeferredPaintPending As Boolean

Private mDeferredPaintAction As DeferredPaintActions

Private mAreasToRedraw() As TRectangle
Private mAreasToRedrawIndex As Long


'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' IDeferredAction Interface Members
'================================================================================

Private Sub IDeferredAction_run(ByVal data As Variant)
Const ProcName As String = "IDeferredAction_run"
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

If mRegion Is Nothing Then Exit Sub

Static et As New ElapsedTimer
If gIsInDev Then
    et.StartTiming
End If

Dim lDeferredPaintAction As DeferredPaintActions
lDeferredPaintAction = mDeferredPaintAction
mDeferredPaintAction = DeferredPaintRepaintNone
mDeferredPaintPending = False

Dim s As String

Select Case lDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintPaintAreas
    s = "Paint areas"
    mRegion.PaintAreas mAreasToRedraw, mAreasToRedrawIndex + 1
Case DeferredPaintPaint
    s = "Paint sync"
    mRegion.PaintSync
Case DeferredPaintChangeGaugeAndPaint
    s = "Change gauge and paint"
    mRegion.ChangeGaugeAndPaint
Case DeferredPaintScaleAndPaint
    s = "Scale and paint"
    mRegion.ScaleAndPaintSync
Case DeferredPaintCalculateAndPaint
    s = "Calculate and paint"
    mRegion.CalculateAndPaintSync
End Select


If gIsInDev Then
    Debug.Print "Deferred paint (" & s & "): " & Format(et.ElapsedTimeMicroseconds / 1000, "0.000ms")
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddAreaToRedraw( _
                ByRef areaToRedraw As TRectangle)
Const ProcName As String = "AddareaToRedraw"
On Error GoTo Err

Dim handled As Boolean

Dim i As Long
For i = 0 To mAreasToRedrawIndex
    If RectContainsRect(mAreasToRedraw(i), areaToRedraw) Then
        handled = True
        Exit For
    ElseIf RectContainsRect(areaToRedraw, mAreasToRedraw(i)) Then
        If Not handled Then
            mAreasToRedraw(i) = areaToRedraw
            handled = True
        Else
            mAreasToRedraw(i).isValid = False
        End If
    End If
Next

If Not handled Then
    mAreasToRedrawIndex = mAreasToRedrawIndex + 1
    If mAreasToRedrawIndex > UBound(mAreasToRedraw) Then
        ReDim Preserve mAreasToRedraw(2 * (UBound(mAreasToRedraw) + 1) - 1) As TRectangle
    End If
    mAreasToRedraw(mAreasToRedrawIndex) = areaToRedraw
End If

doDeferredPaint DeferredPaintPaintAreas

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub CalculateAndPaint()
Const ProcName As String = "CalculateAndPaint"
On Error GoTo Err

doDeferredPaint DeferredPaintCalculateAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ChangeGaugeAndPaint()
Const ProcName As String = "ChangeGaugeAndPaint"
On Error GoTo Err

doDeferredPaint DeferredPaintChangeGaugeAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mRegion = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pRegion As ChartRegion)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mRegion = pRegion
doDeferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Paint()
Const ProcName As String = "Paint"
On Error GoTo Err

doDeferredPaint DeferredPaintPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ScaleAndPaint()
Const ProcName As String = "ScaleAndPaint"
On Error GoTo Err

doDeferredPaint DeferredPaintScaleAndPaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doDeferredPaint(ByVal action As DeferredPaintActions)
Const ProcName As String = "DeferredPaint"
On Error GoTo Err

If action = DeferredPaintRepaintNone Then
    mDeferredPaintAction = DeferredPaintRepaintNone
    mDeferredPaintPending = False
Else
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        mDeferredPaintAction = action
    Case DeferredPaintPaintAreas
        If action = DeferredPaintPaint Or _
            action = DeferredPaintChangeGaugeAndPaint Or _
            action = DeferredPaintScaleAndPaint Or _
            action = DeferredPaintCalculateAndPaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintPaint
        If action = DeferredPaintChangeGaugeAndPaint Or _
            action = DeferredPaintScaleAndPaint Or _
            action = DeferredPaintCalculateAndPaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintChangeGaugeAndPaint
        If action = DeferredPaintScaleAndPaint Or _
            action = DeferredPaintCalculateAndPaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintScaleAndPaint
        If action = DeferredPaintCalculateAndPaint Then mDeferredPaintAction = action
    Case DeferredPaintCalculateAndPaint
        
    End Select
    If Not mDeferredPaintPending Then
        mDeferredPaintPending = True
        DeferAction Me, Empty, 20, ExpiryTimeUnitMilliseconds
    End If
End If

If action <> DeferredPaintPaintAreas Then
    ReDim mAreasToRedraw(3) As TRectangle
    mAreasToRedrawIndex = -1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

