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

Implements IConfigurable

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

'@================================================================================
' IConfigurable Interface Members
'@================================================================================

Private Property Let IConfigurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IConfigurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IConfigurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IConfigurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConfigurable_RemoveFromConfig()
Const ProcName As String = "IConfigurable_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get CoordinateSystem() As CoordinateSystems
CoordinateSystem = mCoordSystem
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"

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
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get Length() As Double
Length = mLength
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aDim As Dimension) As Boolean
Const ProcName As String = "Equals"

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
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pLength As Double, _
                ByVal pCoordSystem As CoordinateSystems)
mLength = pLength
mCoordSystem = pCoordSystem
End Sub

Public Function LengthLogicalX( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "LengthLogicalX"

On Error GoTo Err

Select Case mCoordSystem
Case CoordsLogical
    LengthLogicalX = mLength
Case CoordsRelative
    LengthLogicalX = pRegion.ViewPort.ConvertRelativeToLogicalX(mLength)
Case CoordsDistance
    LengthLogicalX = pRegion.ViewPort.ConvertDistanceToLogicalX(mLength)
Case CoordsCounterDistance
    LengthLogicalX = pRegion.ViewPort.ConvertCounterDistanceToLogicalY(mLength)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LengthLogicalY( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "LengthLogicalY"

On Error GoTo Err

Select Case mCoordSystem
Case CoordsLogical
    LengthLogicalY = mLength
Case CoordsRelative
    LengthLogicalY = pRegion.ViewPort.ConvertRelativeToLogicalY(mLength)
Case CoordsDistance
    LengthLogicalY = pRegion.ViewPort.ConvertDistanceToLogicalY(mLength)
Case CoordsCounterDistance
    LengthLogicalY = pRegion.ViewPort.ConvertCounterDistanceToLogicalY(mLength)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig
mLength = mConfig.GetSetting(ConfigSettingLength, "0")
mCoordSystem = mConfig.GetSetting(ConfigSettingCoordSystem, "0")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' Helper Functions
'================================================================================





