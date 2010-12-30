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

Implements DeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum DeferredPaintActions
    DeferredPaintRepaintNone
    DeferredPaintRepaintChangedAreas
    DeferredPaintRepaintRegion
    DeferredPaintChangeGaugeandRepaint
    DeferredPaintScaleAndRepaint
    DeferredPaintCalculateAndRepaint
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

Private mRegionRef As WeakReference

Private mDeferredPaintPending As Boolean

Private mDeferredPaintAction As DeferredPaintActions

Private mAreasToRedraw() As TRectangle
Private mAreasToRedrawIndex As Long

'Private mAreasToUndraw() As TRectangle
'Private mAreasToUndrawIndex As Long
'
'Private mRedrawObjects() As RedrawObjectsTableEntry
'Private mRedrawObjectsIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal data As Variant)
Dim et As ElapsedTimer
Dim i As Long

Const ProcName As String = "DeferredAction_run"
Dim failpoint As String
On Error GoTo Err

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
#End If

If regionObj.IsCleared Then Exit Sub

If gIsInDev Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

Select Case mDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintRepaintChangedAreas

    failpoint = 50
    
    regionObj.RepaintAreas mAreasToRedraw, mAreasToRedrawIndex + 1

Case DeferredPaintRepaintRegion
    
    failpoint = 300
    
    regionObj.RepaintRegion
Case DeferredPaintChangeGaugeandRepaint
    
    failpoint = 400
    
    regionObj.ChangeGaugeAndRepaint
    
Case DeferredPaintScaleAndRepaint
    
    failpoint = 1000
    
    regionObj.ScaleAndRepaint
    
Case DeferredPaintCalculateAndRepaint
    
    failpoint = 1200
    
    regionObj.CalculateAndRepaint
    
End Select


If gIsInDev Then
    Dim s As String
    s = "Deferred paint: " & Format(et.ElapsedTimeMicroseconds, "0")
    Debug.Print s
End If

doDeferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

'Friend Sub AddAreaToUndraw( _
'                ByRef areaToUndraw As TRectangle)
'Const ProcName As String = "AddAreaToUndraw"
'Dim failpoint As String
'On Error GoTo Err
'
'Dim i As Long
'Dim handled As Boolean
'For i = 0 To mAreasToUndrawIndex
'    If RectOverlaps(areaToUndraw, mAreasToUndraw(i)) Then
'         mAreasToUndraw(i) = RectUnion(areaToUndraw, mAreasToUndraw(i))
'         handled = True
'    End If
'Next
'If Not handled Then
'    mAreasToUndrawIndex = mAreasToUndrawIndex + 1
'    If mAreasToUndrawIndex > UBound(mAreasToUndraw) Then
'        ReDim Preserve mAreasToUndraw(2 * (UBound(mAreasToUndraw) + 1) - 1) As TRectangle
'    End If
'    mAreasToUndraw(mAreasToUndrawIndex) = areaToUndraw
'End If
'doDeferredPaint DeferredPaintRepaintChangedAreas
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'
'End Sub

Friend Sub AddAreaToRedraw( _
                ByRef areaToRedraw As TRectangle)
Dim i As Long
Dim handled As Boolean

Const ProcName As String = "AddareaToRedraw"
Dim failpoint As String
On Error GoTo Err

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

doDeferredPaint DeferredPaintRepaintChangedAreas

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

'Friend Sub AddChangedObject( _
'                ByVal graphObj As IGraphicObject, _
'                ByRef areaToRedraw As TRectangle)
'Dim i As Long
'Dim handled As Boolean
'Dim currRedrawArea As TRectangle
'
'Const ProcName As String = "AddChangedObject"
'Dim failpoint As String
'On Error GoTo Err
'
'For i = 0 To mRedrawObjectsIndex
'    If mRedrawObjects(i).graphObj Is graphObj Then
'        currRedrawArea = mRedrawObjects(i).areaToRedraw
'        mRedrawObjects(i).areaToRedraw = RectUnion(areaToRedraw, mRedrawObjects(i).areaToRedraw)
'        If mRedrawObjects(i).areaToRedraw.isValid Then
'            handled = True
'        Else
'            mRedrawObjects(i).areaToRedraw = currRedrawArea
'        End If
'    End If
'Next
'
'If Not handled Then
'    mRedrawObjectsIndex = mRedrawObjectsIndex + 1
'    If mRedrawObjectsIndex > UBound(mRedrawObjects) Then
'        ReDim Preserve mRedrawObjects(2 * (UBound(mRedrawObjects) + 1) - 1) As RedrawObjectsTableEntry
'    End If
'    Set mRedrawObjects(mRedrawObjectsIndex).graphObj = graphObj
'    mRedrawObjects(mRedrawObjectsIndex).areaToRedraw = areaToRedraw
'End If
'
'doDeferredPaint DeferredPaintRepaintChangedAreas
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
'End Sub

Friend Sub CalculateAndRepaint()
Const ProcName As String = "CalculateAndRepaint"
Dim failpoint As String
On Error GoTo Err

doDeferredPaint DeferredPaintCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub ChangeGaugeAndRepaint()
Const ProcName As String = "ChangeGaugeAndRepaint"
Dim failpoint As String
On Error GoTo Err

doDeferredPaint DeferredPaintChangeGaugeandRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pRegion As ChartRegion)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mRegionRef = CreateWeakReference(pRegion)
doDeferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RepaintRegion()
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
On Error GoTo Err

doDeferredPaint DeferredPaintRepaintRegion

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub ScaleAndRepaint()
Const ProcName As String = "ScaleAndRepaint"
Dim failpoint As String
On Error GoTo Err

doDeferredPaint DeferredPaintScaleAndRepaint

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doDeferredPaint(ByVal action As DeferredPaintActions)
Const ProcName As String = "DeferredPaint"
Dim failpoint As String
On Error GoTo Err

If action = DeferredPaintRepaintNone Then
    mDeferredPaintAction = DeferredPaintRepaintNone
    mDeferredPaintPending = False
Else
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        mDeferredPaintAction = action
    Case DeferredPaintRepaintChangedAreas
        If action = DeferredPaintRepaintRegion Or _
            action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintRepaintRegion
        If action = DeferredPaintChangeGaugeandRepaint Or _
            action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintChangeGaugeandRepaint
        If action = DeferredPaintScaleAndRepaint Or _
            action = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = action
    Case DeferredPaintScaleAndRepaint
        If action = DeferredPaintCalculateAndRepaint Then mDeferredPaintAction = action
    Case DeferredPaintCalculateAndRepaint
        
    End Select
    If Not mDeferredPaintPending Then
        mDeferredPaintPending = True
        DeferAction Me, Empty, 20, ExpiryTimeUnitMilliseconds
    End If
End If

If action <> DeferredPaintRepaintChangedAreas Then
    ReDim mAreasToRedraw(3) As TRectangle
    mAreasToRedrawIndex = -1
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Function regionObj() As ChartRegion
Const ProcName As String = "regionObj"
Dim failpoint As String
On Error GoTo Err

Set regionObj = mRegionRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function
