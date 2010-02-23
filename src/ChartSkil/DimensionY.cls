VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DimensionY"
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

Private Const ModuleName                As String = "DimensionY"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mY As Double
Private mViewport As Viewport

Private mCoordSystemY As CoordinateSystems

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

Public Property Get Y() As Double
Y = mY
End Property

Public Property Get yLogical() As Double
Const ProcName As String = "YLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystemY
Case CoordsLogical
    yLogical = mY
Case CoordsRelative
    yLogical = mViewport.ConvertRelativeToLogicalY(mY)
Case CoordsDistance
    yLogical = mViewport.ConvertDistanceToLogicalY(mY)
Case CoordsCounterDistance
    yLogical = mViewport.ConvertCounterDistanceToLogicalY(mY)
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
    If .Y <> mY Then Exit Function
End With
Equals = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal Y As Double, _
                ByVal coordSystemY As CoordinateSystems, _
                ByVal pViewport As Viewport)
mY = Y
mCoordSystemY = coordSystemY
Set mViewport = pViewport
End Sub

'================================================================================
' Helper Functions
'================================================================================





