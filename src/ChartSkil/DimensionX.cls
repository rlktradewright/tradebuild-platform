VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Dimension"
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

Private Const ModuleName                As String = "Dimension"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mLength As Double

Private mCoordSystem As CoordinateSystems

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get Length() As Double
Length = mLength
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
If aDim.Length <> mLength Then Exit Function
Equals = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal pLength As Double, _
                ByVal pCoordSystem As CoordinateSystems)
mLength = Length
mCoordSystem = pCoordSystem
End Sub

Public Function LengthLogicalX( _
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "LengthLogicalX"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystem
Case CoordsLogical
    LengthLogicalX = mLength
Case CoordsRelative
    LengthLogicalX = pViewport.ConvertRelativeToLogicalX(mLength)
Case CoordsDistance
    LengthLogicalX = pViewport.ConvertDistanceToLogicalX(mLength)
Case CoordsCounterDistance
    LengthLogicalX = pViewport.ConvertCounterDistanceToLogicalY(mLength)
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Public Function LengthLogicalY( _
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "LengthLogicalY"
Dim failpoint As String
On Error GoTo Err

Select Case mCoordSystem
Case CoordsLogical
    LengthLogicalY = mLength
Case CoordsRelative
    LengthLogicalY = pViewport.ConvertRelativeToLogicalY(mLength)
Case CoordsDistance
    LengthLogicalY = pViewport.ConvertDistanceToLogicalY(mLength)
Case CoordsCounterDistance
    LengthLogicalY = pViewport.ConvertCounterDistanceToLogicalY(mLength)
End Select

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'================================================================================
' Helper Functions
'================================================================================





