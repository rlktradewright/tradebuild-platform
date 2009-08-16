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
' Member variables and constants
'================================================================================

Private mRegionRef As WeakReference

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

Friend Property Let Region( _
                ByVal value As ChartRegion)
Set mRegionRef = CreateWeakReference(value)
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub AddGraphicObject( _
                ByVal graphObj As IGraphicObject)
regionObj.GOSP_AddGraphicObject graphObj
End Sub

Friend Sub Finish()
Set mRegionRef = Nothing
End Sub

Friend Sub ObjectChanged(ByVal Handle As Long, _
                ByRef areaToRedraw As TRectangle)
regionObj.GOSP_ObjectChanged Handle, areaToRedraw
End Sub

Friend Sub ObjectUndrawn(ByVal Handle As Long, _
                ByRef undrawnArea As TRectangle)
regionObj.GOSP_ObjectUndrawn Handle, undrawnArea
End Sub

Friend Sub PaintRegion( _
                ByVal recalculate As Boolean)
regionObj.GOSP_PaintRegion recalculate
End Sub

Friend Sub RecalcBoundary( _
                ByVal SeriesID As Long)
regionObj.GOSP_RecalcBoundary SeriesID
End Sub

Friend Sub RemoveObject( _
                ByVal Handle As Long)
regionObj.GOSP_RemoveObject Handle
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function
