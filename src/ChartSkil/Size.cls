VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Size"
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

Private Const ModuleName                As String = "Size"

Private Const ConfigSettingCoordSystemHeight       As String = "&CoordSystemHeight"
Private Const ConfigSettingCoordSystemWidth        As String = "&CoordSystemWidth"
Private Const ConfigSettingHeight       As String = "&Height"
Private Const ConfigSettingWidth        As String = "&Width"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mWidth As Double
Private mHeight As Double

Private mCoordSystemWidth As CoordinateSystems
Private mCoordSystemHeight As CoordinateSystems

Private mConfig As ConfigurationSection

'================================================================================
' Class Event Handlers
'================================================================================

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
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

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


mConfig.SetSetting ConfigSettingWidth, mWidth
mConfig.SetSetting ConfigSettingHeight, mHeight
mConfig.SetSetting ConfigSettingCoordSystemWidth, mCoordSystemWidth
mConfig.SetSetting ConfigSettingCoordSystemHeight, mCoordSystemHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get CoordinateSystemHeight() As CoordinateSystems
CoordinateSystemHeight = mCoordSystemHeight
End Property

Public Property Get CoordinateSystemWidth() As CoordinateSystems
CoordinateSystemWidth = mCoordSystemWidth
End Property

Public Property Get Width() As Double
Width = mWidth
End Property

Public Property Get Height() As Double
Height = mHeight
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aDim As Size) As Boolean
Const ProcName As String = "Equals"

On Error GoTo Err

If aDim Is Nothing Then Exit Function
If aDim Is Me Then
    Equals = True
    Exit Function
End If
With aDim
    If .CoordinateSystemWidth <> mCoordSystemWidth Then Exit Function
    If .CoordinateSystemHeight <> mCoordSystemHeight Then Exit Function
    If .Width <> mWidth Then Exit Function
    If .Height <> mHeight Then Exit Function
End With
Equals = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pWidth As Double, _
                ByVal pHeight As Double, _
                ByVal pCoordSystemWidth As CoordinateSystems, _
                ByVal pCoordSystemHeight As CoordinateSystems)
mWidth = pWidth
mHeight = pHeight
mCoordSystemWidth = pCoordSystemWidth
mCoordSystemHeight = pCoordSystemHeight
End Sub

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig
mHeight = mConfig.GetSetting(ConfigSettingHeight, "0")
mWidth = mConfig.GetSetting(ConfigSettingWidth, "0")
mCoordSystemHeight = mConfig.GetSetting(ConfigSettingCoordSystemHeight, "0")
mCoordSystemWidth = mConfig.GetSetting(ConfigSettingCoordSystemWidth, "0")

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

Public Function WidthLogical( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "WidthLogical"

On Error GoTo Err

Select Case mCoordSystemWidth
Case CoordsLogical
    WidthLogical = mWidth
Case CoordsRelative
    WidthLogical = pRegion.ViewPort.ConvertRelativeToLogicalX(mWidth)
Case CoordsDistance
    WidthLogical = pRegion.ViewPort.ConvertDistanceToLogicalX(mWidth)
Case CoordsCounterDistance
    WidthLogical = pRegion.ViewPort.ConvertCounterDistanceToLogicalX(mWidth)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HeightLogical( _
                ByVal pRegion As ChartRegion) As Double
Const ProcName As String = "HeightLogical"

On Error GoTo Err

Select Case mCoordSystemHeight
Case CoordsLogical
    HeightLogical = mHeight
Case CoordsRelative
    HeightLogical = pRegion.ViewPort.ConvertRelativeToLogicalY(mHeight)
Case CoordsDistance
    HeightLogical = pRegion.ViewPort.ConvertDistanceToLogicalY(mHeight)
Case CoordsCounterDistance
    HeightLogical = pRegion.ViewPort.ConvertCounterDistanceToLogicalY(mHeight)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================



