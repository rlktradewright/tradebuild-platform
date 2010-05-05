VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DimensionX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                As String = "DimensionX"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mX As Double
Private mViewport As Viewport

Private mCoordSystemX As CoordinateSystems

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get Viewport() As Viewport
Set Viewport = mViewport
End Property

Public Property Get X() As Double
X = mX
End Property

Public Property Get XLogical() As Double
Const ProcName As String = "XLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystemX
Case CoordsLogical
    XLogical = mX
Case CoordsRelative
    XLogical = mViewport.ConvertRelativeToLogicalX(mX)
Case CoordsDistance
    XLogical = mViewport.ConvertDistanceToLogicalX(mX)
Case CoordsCounterDistance
    XLogical = mViewport.ConvertCounterDistanceToLogicalY(mX)
End Select

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aDim As Dimension) As Boolean
Const ProcName As String = "Equals"
Dim failpoint As String
On Error GoTo Err

If aDim Is Nothing Then Exit Function
If aDim Is Me Then
    Equals = True
    Exit Function
End If
With aDim
    If Not .Viewport Is mViewport Then Exit Function
    If .X <> mX Then Exit Function
End With
Equals = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal X As Double, _
                ByVal coordSystemX As CoordinateSystems, _
                ByVal pViewport As Viewport)
mX = X
mCoordSystemX = coordSystemX
Set mViewport = pViewport
End Sub

'================================================================================
' Helper Functions
'================================================================================




