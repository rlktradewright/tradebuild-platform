Attribute VB_Name = "GPoint"
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

Private Const ModuleName                            As String = "GPoint"

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

Public Function gLoadDimensionFromConfig( _
                ByVal pConfig As ConfigurationSection) As Dimension
Const ProcName As String = "gLoadDimensionFromConfig"
On Error GoTo Err

Set gLoadDimensionFromConfig = New Dimension
gLoadDimensionFromConfig.LoadFromConfig pConfig

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gLoadSizeFromConfig( _
                ByVal pConfig As ConfigurationSection) As Size
Const ProcName As String = "gLoadSizeFromConfig"
On Error GoTo Err

Set gLoadSizeFromConfig = New Size
gLoadSizeFromConfig.LoadFromConfig pConfig

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gNewDimension( _
                ByVal pLength As Double, _
                Optional ByVal pCoordSystem As CoordinateSystems = CoordsDistance) As Dimension
Const ProcName As String = "gNewDimension"
Dim failpoint As String
On Error GoTo Err

Set gNewDimension = New Dimension
gNewDimension.Initialise pLength, pCoordSystem
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gNewPoint( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical, _
                Optional ByVal Offset As Size) As Point
Const ProcName As String = "gNewPoint"
Dim failpoint As String
On Error GoTo Err

Set gNewPoint = New Point
gNewPoint.Initialise X, Y, coordSystemX, coordSystemY, Offset

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gNewSize( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsDistance, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsDistance) As Size
Const ProcName As String = "gNewSize"
Dim failpoint As String
On Error GoTo Err

Set gNewSize = New Size
gNewSize.Initialise X, Y, coordSystemX, coordSystemY
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTransformCoordX( _
                ByVal pValue As Double, _
                ByVal pFromCoordSys As CoordinateSystems, _
                ByVal pToCoordSys As CoordinateSystems, _
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "gTransformCoordX"
Dim failpoint As String
On Error GoTo Err

If pFromCoordSys = pToCoordSys Then
    gTransformCoordX = pValue
    Exit Function
End If

Select Case pFromCoordSys
Case CoordsLogical
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pViewport.ConvertLogicalToCounterDistanceX(pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pViewport.ConvertLogicalToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pViewport.ConvertLogicalToRelativeX(pValue)
    End If
Case CoordsRelative
    If pToCoordSys = CoordsLogical Then
        gTransformCoordX = pViewport.ConvertRelativeToLogicalX(pValue) + pViewport.Boundary.Left
    ElseIf pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pViewport.ConvertRelativeToCounterDistanceX(pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pViewport.ConvertRelativeToDistanceX(pValue)
    End If
Case CoordsDistance
    If pToCoordSys = CoordsLogical Then
        gTransformCoordX = pViewport.ConvertDistanceToLogicalX(pValue) + pViewport.Boundary.Left
    ElseIf pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pViewport.ConvertDistanceToCounterDistanceX(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pViewport.ConvertDistanceToRelativeX(pValue)
    End If
Case CoordsCounterDistance
    If pToCoordSys = CoordsLogical Then
        gTransformCoordX = pViewport.ConvertCounterDistanceToLogicalY(pValue) + pViewport.Boundary.Left
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pViewport.ConvertCounterDistanceToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pViewport.ConvertCounterDistanceToRelativeX(pValue)
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function gTransformCoordY( _
                ByVal pValue As Double, _
                ByVal pFromCoordSys As CoordinateSystems, _
                ByVal pToCoordSys As CoordinateSystems, _
                ByVal pViewport As Viewport) As Double
Const ProcName As String = "gTransformCoordY"
Dim failpoint As String
On Error GoTo Err

If pFromCoordSys = pToCoordSys Then
    gTransformCoordY = pValue
    Exit Function
End If

Select Case pFromCoordSys
Case CoordsLogical
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pViewport.ConvertLogicalToCounterDistanceY(pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pViewport.ConvertLogicalToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pViewport.ConvertLogicalToRelativeY(pValue)
    End If
Case CoordsRelative
    If pToCoordSys = CoordsLogical Then
        gTransformCoordY = pViewport.ConvertRelativeToLogicalY(pValue) + pViewport.Boundary.Bottom
    ElseIf pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pViewport.ConvertRelativeToCounterDistanceY(pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pViewport.ConvertRelativeToDistanceY(pValue)
    End If
Case CoordsDistance
    If pToCoordSys = CoordsLogical Then
        gTransformCoordY = pViewport.ConvertDistanceToLogicalY(pValue) + pViewport.Boundary.Bottom
    ElseIf pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pViewport.ConvertDistanceToCounterDistanceY(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pViewport.ConvertDistanceToRelativeY(pValue)
    End If
Case CoordsCounterDistance
    If pToCoordSys = CoordsLogical Then
        gTransformCoordY = pViewport.ConvertCounterDistanceToLogicalY(pValue) + pViewport.Boundary.Bottom
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pViewport.ConvertCounterDistanceToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pViewport.ConvertCounterDistanceToRelativeY(pValue)
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


