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

Private mRegion As ChartRegion

'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let Region( _
                ByVal value As ChartRegion)
Set mRegion = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub AddGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long)
mRegion.GOSP_AddGraphicObject graphObj, seriesId
End Sub

Friend Sub ObjectChanged(ByVal Handle As Long, _
                ByRef areaToRedraw As TRectangle)
mRegion.GOSP_ObjectChanged Handle, areaToRedraw
End Sub

Friend Sub ObjectUndrawn(ByVal Handle As Long, _
                ByRef undrawnArea As TRectangle)
mRegion.GOSP_ObjectUndrawn Handle, undrawnArea
End Sub

Friend Sub PaintRegion( _
                ByVal recalculate As Boolean)
mRegion.GOSP_PaintRegion recalculate
End Sub

Friend Sub RecalcBoundary( _
                ByVal seriesId As Long)
mRegion.GOSP_RecalcBoundary seriesId
End Sub

Friend Sub RemoveObject( _
                ByVal Handle As Long)
mRegion.GOSP_RemoveObject Handle
End Sub

'================================================================================
' Helper Functions
'================================================================================


