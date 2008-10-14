VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartRegionEventProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

Event addGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long)

Event objectChanged( _
                ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)

Event objectUndrawn( _
                ByVal handle As Long, _
                ByRef undrawnArea As TRectangle)

Event paintRegion( _
                ByVal recalculate As Boolean)

Event recalcBoundary( _
                ByVal seriesId As Long)

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

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

'================================================================================
' Methods
'================================================================================

Friend Sub addGraphicObject( _
                ByVal graphObj As IGraphicObject, _
                ByVal seriesId As Long)
RaiseEvent addGraphicObject(graphObj, seriesId)
End Sub

Friend Sub objectChanged(ByVal handle As Long, _
                    ByRef areaToRedraw As TRectangle)
RaiseEvent objectChanged(handle, areaToRedraw)
End Sub

Friend Sub objectUndrawn(ByVal handle As Long, _
                    ByRef undrawnArea As TRectangle)
RaiseEvent objectUndrawn(handle, undrawnArea)
End Sub

Friend Sub paintRegion( _
                ByVal recalculate As Boolean)
RaiseEvent paintRegion(recalculate)
End Sub

Friend Sub recalcBoundary( _
                ByVal seriesId As Long)
RaiseEvent recalcBoundary(seriesId)
End Sub

'================================================================================
' Helper Functions
'================================================================================


