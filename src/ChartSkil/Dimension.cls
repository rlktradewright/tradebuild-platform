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

Private Const ConfigSettingCoordSystem  As String = "&CoordSystem"
Private Const ConfigSettingLength       As String = "&Length"

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

Private mConfig As ConfigurationSection

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = Value


mConfig.SetSetting ConfigSettingLength, mLength
mConfig.SetSetting ConfigSettingCoordSystem, mCoordSystem

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

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

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
mLength = mConfig.GetSetting(ConfigSettingLength, "0")
mCoordSystem = mConfig.GetSetting(ConfigSettingCoordSystem, "0")

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'================================================================================
' Helper Functions
'================================================================================





