VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "GraphObjServiceProvider"

'================================================================================
' Member variables
'================================================================================

Private mRegionRef                          As WeakReference
Private mData                               As ChartRegionData

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Terminate()
Debug.Print "GraphObjServiceProvider terminated"
End Sub

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub AddGraphicObject( _
                ByVal graphObj As IGraphicObject)
Const ProcName As String = "AddGraphicObject"

On Error GoTo Err

mData.AddGraphicObject graphObj

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Set mRegionRef = Nothing
End Sub

Friend Sub Initialise(ByVal pRegion As ChartRegion, ByVal pData As ChartRegionData)
Set mRegionRef = CreateWeakReference(pRegion)
Set mData = pData
End Sub

Friend Sub ObjectChanged(ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Const ProcName As String = "ObjectChanged"
On Error GoTo Err

Dim graphObj As IGraphicObject
Dim RedrawArea As TRectangle

Set graphObj = mData.GetGraphicObject(handle)
If Not mData.UpdateGraphicObject(graphObj) Then Exit Sub

If Not regionObj.IsDrawingEnabled Then Exit Sub

If Not graphObj.BoundingRectangle.isValid Then Exit Sub
If graphObj.NoDraw Then Exit Sub

If regionObj.PaintingInProgress Then Exit Sub

RedrawArea = RectIntersection(areaToRedraw, regionObj.ViewPort.Boundary)

' check if the vertical bounds have changed, and redraw if so
If regionObj.Autoscaling And graphObj.IncludeInAutoscale Then
    Dim verticalRangeInViewport As TInterval
    verticalRangeInViewport = graphObj.VerticalRangeInXinterval(RectGetXInterval(regionObj.ViewPort.Boundary))
    
    If verticalRangeInViewport.isValid And _
        regionObj.Autoscaling And _
        (verticalRangeInViewport.endValue > regionObj.ViewPort.Boundary.Top Or _
            verticalRangeInViewport.startValue < (regionObj.ViewPort.Boundary.Bottom)) Then
        regionObj.ScaleAndPaint
        Exit Sub
    End If
End If

regionObj.RedrawObject graphObj, RedrawArea

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ObjectUndrawn(ByVal handle As Long, _
                ByRef undrawnArea As TRectangle)
Const ProcName As String = "ObjectUndrawn"

On Error GoTo Err

regionObj.UndrawArea undrawnArea, mData.GetGraphicObject(handle)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintRegion( _
                ByVal recalculate As Boolean)
Const ProcName As String = "PaintRegion"

On Error GoTo Err

If recalculate Then
    regionObj.CalculateAndPaint
Else
    regionObj.Paint
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RecalcBoundary( _
                ByVal SeriesID As Long)
Const ProcName As String = "RecalcBoundary"

On Error GoTo Err

mData.RecalcBoundaryForObjectsInSeries SeriesID

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveObject( _
                ByVal handle As Long)
Const ProcName As String = "RemoveObject"

On Error GoTo Err

regionObj.RemoveGraphicObjectEx mData.GetGraphicObject(handle), True, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function
