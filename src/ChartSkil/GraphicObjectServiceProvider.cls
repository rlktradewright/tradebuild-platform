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
Const ProcName As String = "AddGraphicObject"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_AddGraphicObject graphObj

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub Finish()
Set mRegionRef = Nothing
End Sub

Friend Sub ObjectChanged(ByVal handle As Long, _
                ByRef areaToRedraw As TRectangle)
Const ProcName As String = "ObjectChanged"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_ObjectChanged handle, areaToRedraw

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub ObjectUndrawn(ByVal handle As Long, _
                ByRef undrawnArea As TRectangle)
Const ProcName As String = "ObjectUndrawn"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_ObjectUndrawn handle, undrawnArea

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub PaintRegion( _
                ByVal recalculate As Boolean)
Const ProcName As String = "PaintRegion"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_PaintRegion recalculate

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RecalcBoundary( _
                ByVal SeriesID As Long)
Const ProcName As String = "RecalcBoundary"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_RecalcBoundary SeriesID

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub RemoveObject( _
                ByVal handle As Long)
Const ProcName As String = "RemoveObject"
Dim failpoint As String
On Error GoTo Err

regionObj.GOSP_RemoveObject handle

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function regionObj() As ChartRegion
Set regionObj = mRegionRef.Target
End Function
