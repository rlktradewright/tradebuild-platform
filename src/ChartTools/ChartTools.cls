VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartTools"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Private Const ProjectName                   As String = "ChartTools26"
Private Const ModuleName                    As String = "ChartTools"

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

'@================================================================================
' Methods
'@================================================================================

Public Function CreateFibRetracementTool( _
                ByVal controller As ChartController, _
                ByRef pLineSpecs() As FibLineSpecifier, _
                Optional ByVal layer As LayerNumbers = -1) As FibRetracementTool
If controller Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateLineTool", _
            "You must supply a ChartController"
End If

Set CreateFibRetracementTool = New FibRetracementTool
CreateFibRetracementTool.Initialise controller, pLineSpecs, layer
End Function

Public Function CreateLineTool( _
                ByVal controller As ChartController, _
                Optional ByVal pLineStyle As LineStyle, _
                Optional ByVal layer As LayerNumbers = -1) As LineTool
If controller Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "CreateLineTool", _
            "You must supply a ChartController"
End If

Set CreateLineTool = New LineTool
CreateLineTool.Initialise controller, pLineStyle, layer
End Function

'@================================================================================
' Helper Functions
'@================================================================================


