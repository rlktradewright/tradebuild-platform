VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IChartTool"
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


Private Const ModuleName                    As String = "IChartTool"

'@================================================================================
' Member variables
'@================================================================================

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

Public Property Let Region( _
                ByVal value As ChartRegion)
                
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Click( _
                ByVal X As Double, _
                ByVal Y As Double)

End Sub

Public Sub Create( _
                ByVal X As Double, _
                ByVal Y As Double)

End Sub

Public Sub DblClick()

End Sub

Public Sub Delete()

End Sub

Public Sub EndResize()

End Sub

Public Sub Deselected()

End Sub

Public Sub Move( _
                ByVal X As Double, _
                ByVal Y As Double)

End Sub

Public Sub MoveResize( _
                ByVal selHandleIndex As Long, _
                ByVal X As Double, _
                ByVal Y As Double)

End Sub

Public Sub ResizeCreate( _
                ByVal X As Double, _
                ByVal Y As Double)

End Sub

Public Sub StartMove( _
                ByVal X As Single, _
                ByVal Y As Single)

End Sub

Public Sub StartResize( _
                ByVal selHandleIndex As Long, _
                ByVal X As Single, _
                ByVal Y As Single)

End Sub

Public Sub Undo()

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


