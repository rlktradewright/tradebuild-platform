VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IGraphicObjectStyle

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEventData)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "LineStyle"

Private Const ConfigSettingColor                As String = "&Color"
Private Const ConfigSettingThickness            As String = "&Thickness"
Private Const ConfigSettingLineStyle            As String = "&LineStyle"
Private Const ConfigSettingExtendBefore         As String = "&ExtendBefore"
Private Const ConfigSettingExtendAfter          As String = "&ExtendAfter"
Private Const ConfigSettingArrowStartStyle      As String = "&ArrowStartStyle"
Private Const ConfigSettingArrowStartLength     As String = "&ArrowStartLength"
Private Const ConfigSettingArrowStartWidth      As String = "&ArrowStartWidth"
Private Const ConfigSettingArrowStartColor      As String = "&ArrowStartColor"
Private Const ConfigSettingArrowStartFillColor  As String = "&ArrowStartFillColor"
Private Const ConfigSettingArrowStartfillStyle  As String = "&ArrowStartFillStyle"
Private Const ConfigSettingArrowEndStyle        As String = "&ArrowEndStyle"
Private Const ConfigSettingArrowEndLength       As String = "&ArrowEndLength"
Private Const ConfigSettingArrowEndWidth        As String = "&ArrowEndWidth"
Private Const ConfigSettingArrowEndColor        As String = "&ArrowEndColor"
Private Const ConfigSettingArrowEndFillColor    As String = "&ArrowEndFillColor"
Private Const ConfigSettingArrowEndFillStyle    As String = "&ArrowEndFillStyle"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"
Private Const ConfigSettingFixedX               As String = "&FixedX"
Private Const ConfigSettingFixedY               As String = "&FixedY"
Private Const ConfigSettingExtended             As String = "&Extended"
Private Const ConfigSettingLayer                As String = "&Layer"

Private Const ConfigSectionOffset1              As String = "Offset1"
Private Const ConfigSectionOffset2              As String = "Offset2"

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mStyleType As String
Private mConfig As ConfigurationSection

Private mColor As Long
Private mThickness As Long
Private mLineStyle As LineStyles
Private mExtendBefore As Boolean
Private mExtendAfter As Boolean
Private mArrowStartStyle As ArrowStyles
Private mArrowStartLength As Long
Private mArrowStartWidth As Long
Private mArrowStartColor As Long
Private mArrowStartFillColor As Long
Private mArrowStartfillStyle As FillStyles
Private mArrowEndStyle As ArrowStyles
Private mArrowEndLength As Long
Private mArrowEndWidth As Long
Private mArrowEndColor As Long
Private mArrowEndFillColor As Long
Private mArrowEndFillStyle As FillStyles
Private mIncludeInAutoscale As Boolean
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mExtended As Boolean
Private mOffset1 As Size
Private mOffset2 As Size
Private mLayer As Long

Private mPropertyFlags As LinePropertyFlags

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Line Styles"
End Sub

'@================================================================================
' IGraphicObjectStyle Interface Members
'@================================================================================

Private Property Let IGraphicObjectStyle_ConfigurationSection(ByVal RHS As ConfigurationSection)
Me.ConfigurationSection = RHS
End Property

Private Sub IGraphicObjectStyle_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Me.LoadFromConfig pConfig
End Sub

Private Property Get IGraphicObjectStyle_Name() As String
IGraphicObjectStyle_Name = mName
End Property

Private Property Get IGraphicObjectStyle_ProgId() As String
IGraphicObjectStyle_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IGraphicObjectStyle_StyleType() As String
IGraphicObjectStyle_StyleType = mStyleType
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let ArrowEndColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowEndColor, LinePropertyArrowEndColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndColor() As Long
Const ProcName As String = "ArrowEndColor"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndColor
ArrowEndColor = mArrowEndColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowEndFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowEndFillColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowEndFillColor, LinePropertyArrowEndFillColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndFillColor() As Long
Const ProcName As String = "ArrowEndFillColor"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndFillColor
ArrowEndFillColor = mArrowEndFillColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowEndFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowEndFillStyle"

On Error GoTo Err

setLongProperty mArrowEndFillStyle, LinePropertyArrowEndFillStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
Const ProcName As String = "ArrowEndFillStyle"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndFillStyle
ArrowEndFillStyle = mArrowEndFillStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowEndLength(ByVal Value As Long)
Const ProcName As String = "ArrowEndLength"

On Error GoTo Err

setLongProperty mArrowEndLength, LinePropertyArrowEndLength, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndLength() As Long
Const ProcName As String = "ArrowEndLength"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndLength
ArrowEndLength = mArrowEndLength

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowEndStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowEndStyle"

On Error GoTo Err

setLongProperty mArrowEndStyle, LinePropertyArrowEndStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
Const ProcName As String = "ArrowEndStyle"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndStyle
ArrowEndStyle = mArrowEndStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowEndWidth(ByVal Value As Long)
Const ProcName As String = "ArrowEndWidth"

On Error GoTo Err

setLongProperty mArrowEndWidth, LinePropertyArrowEndWidth, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowEndWidth() As Long
Const ProcName As String = "ArrowEndWidth"
On Error GoTo Err

checkPropertySet LinePropertyArrowEndWidth
ArrowEndWidth = mArrowEndWidth

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowStartColor, LinePropertyArrowStartColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartColor() As Long
Const ProcName As String = "ArrowStartColor"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartColor
ArrowStartColor = mArrowStartColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartFillColor(ByVal Value As Long)
Const ProcName As String = "ArrowStartFillColor"

On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mArrowStartFillColor, LinePropertyArrowStartFillColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartFillColor() As Long
Const ProcName As String = "ArrowStartFillColor"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartFillColor
ArrowStartFillColor = mArrowStartFillColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "ArrowStartFillStyle"

On Error GoTo Err

setLongProperty mArrowStartfillStyle, LinePropertyArrowStartFillStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
Const ProcName As String = "ArrowStartFillStyle"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartFillStyle
ArrowStartFillStyle = mArrowStartfillStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartLength(ByVal Value As Long)
Const ProcName As String = "ArrowStartLength"

On Error GoTo Err

setLongProperty mArrowStartLength, LinePropertyArrowStartLength, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartLength() As Long
Const ProcName As String = "ArrowStartLength"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartLength
ArrowStartLength = mArrowStartLength

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartStyle(ByVal Value As ArrowStyles)
Const ProcName As String = "ArrowStartStyle"

On Error GoTo Err

setLongProperty mArrowStartStyle, LinePropertyArrowStartStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
Const ProcName As String = "ArrowStartStyle"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartStyle
ArrowStartStyle = mArrowStartStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ArrowStartWidth(ByVal Value As Long)
Const ProcName As String = "ArrowStartWidth"

On Error GoTo Err

setLongProperty mArrowStartWidth, LinePropertyArrowStartWidth, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ArrowStartWidth() As Long
Const ProcName As String = "ArrowStartWidth"
On Error GoTo Err

checkPropertySet LinePropertyArrowStartWidth
ArrowStartWidth = mArrowStartWidth

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"

On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mColor, LinePropertyColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
On Error GoTo Err

checkPropertySet LinePropertyColor
Color = mColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If IsPropertySet(LinePropertyArrowEndColor) Then mConfig.SetSetting ConfigSettingArrowEndColor, mArrowEndColor
If IsPropertySet(LinePropertyArrowEndFillColor) Then mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
If IsPropertySet(LinePropertyArrowEndFillStyle) Then mConfig.SetSetting ConfigSettingArrowEndFillStyle, mArrowEndFillStyle
If IsPropertySet(LinePropertyArrowEndLength) Then mConfig.SetSetting ConfigSettingArrowEndLength, mArrowEndLength
If IsPropertySet(LinePropertyArrowEndStyle) Then mConfig.SetSetting ConfigSettingArrowEndStyle, mArrowEndStyle
If IsPropertySet(LinePropertyArrowEndWidth) Then mConfig.SetSetting ConfigSettingArrowEndWidth, mArrowEndWidth

If IsPropertySet(LinePropertyArrowStartColor) Then mConfig.SetSetting ConfigSettingArrowStartColor, mArrowStartColor
If IsPropertySet(LinePropertyArrowStartFillColor) Then mConfig.SetSetting ConfigSettingArrowStartFillColor, mArrowStartFillColor
If IsPropertySet(LinePropertyArrowStartFillStyle) Then mConfig.SetSetting ConfigSettingArrowStartfillStyle, mArrowStartfillStyle
If IsPropertySet(LinePropertyArrowStartLength) Then mConfig.SetSetting ConfigSettingArrowStartLength, mArrowStartLength
If IsPropertySet(LinePropertyArrowStartStyle) Then mConfig.SetSetting ConfigSettingArrowStartStyle, mArrowStartStyle
If IsPropertySet(LinePropertyArrowStartWidth) Then mConfig.SetSetting ConfigSettingArrowStartWidth, mArrowStartWidth

If IsPropertySet(LinePropertyColor) Then mConfig.SetSetting ConfigSettingColor, mColor

If IsPropertySet(LinePropertyExtendAfter) Then mConfig.SetSetting ConfigSettingExtendAfter, mExtendAfter
If IsPropertySet(LinePropertyExtendBefore) Then mConfig.SetSetting ConfigSettingExtendBefore, mExtendBefore
If IsPropertySet(LinePropertyExtended) Then mConfig.SetSetting ConfigSettingExtended, mExtended

If IsPropertySet(LinePropertyFixedX) Then mConfig.SetSetting ConfigSettingFixedX, mFixedX
If IsPropertySet(LinePropertyFixedY) Then mConfig.SetSetting ConfigSettingFixedY, mFixedY

If IsPropertySet(LinePropertyIncludeInAutoscale) Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale

If IsPropertySet(LinePropertyLayer) Then mConfig.SetSetting ConfigSettingLayer, mLayer
If IsPropertySet(LinePropertyLineStyle) Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle

If IsPropertySet(LinePropertyThickness) Then mConfig.SetSetting ConfigSettingThickness, mThickness

If Not mOffset1 Is Nothing Then mOffset1.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionOffset1)
If Not mOffset2 Is Nothing Then mOffset2.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionOffset2)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Let ExtendAfter(ByVal Value As Boolean)
Const ProcName As String = "ExtendAfter"

On Error GoTo Err

setBooleanProperty mExtendAfter, LinePropertyExtendAfter, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ExtendAfter() As Boolean
Const ProcName As String = "ExtendAfter"
On Error GoTo Err

checkPropertySet LinePropertyExtendAfter
ExtendAfter = mExtendAfter

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ExtendBefore(ByVal Value As Boolean)
Const ProcName As String = "ExtendBefore"

On Error GoTo Err

setBooleanProperty mExtendBefore, LinePropertyExtendBefore, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ExtendBefore() As Boolean
Const ProcName As String = "ExtendBefore"
On Error GoTo Err

checkPropertySet LinePropertyExtendBefore
ExtendBefore = mExtendBefore

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"

On Error GoTo Err

setBooleanProperty mExtended, LinePropertyExtended, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
On Error GoTo Err

checkPropertySet LinePropertyExtended
Extended = mExtended

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"

On Error GoTo Err

setBooleanProperty mFixedX, LinePropertyFixedX, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"
On Error GoTo Err

checkPropertySet LinePropertyFixedX
FixedX = mFixedX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"

On Error GoTo Err

setBooleanProperty mFixedY, LinePropertyFixedY, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"
On Error GoTo Err

checkPropertySet LinePropertyFixedY
FixedY = mFixedY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"

On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, LinePropertyIncludeInAutoscale, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

checkPropertySet LinePropertyIncludeInAutoscale
IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"

On Error GoTo Err

If Value < LayerMin Or Value > LayerMax Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid layer number"

setLongProperty mLayer, BarPropertyLayer, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

checkPropertySet BarPropertyLayer
Layer = mLayer

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"

On Error GoTo Err

setLongProperty mLineStyle, LinePropertyLineStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LineStyle() As LineStyles
Const ProcName As String = "LineStyle"
On Error GoTo Err

checkPropertySet LinePropertyLineStyle
LineStyle = mLineStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
On Error GoTo Err

If Value = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Name cannot be blank"
End If

mName = Value
If Not mConfig Is Nothing Then mConfig.InstanceQualifier = mName

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Offset1(ByVal Value As Size)

Const ProcName As String = "Offset1"
On Error GoTo Err

setSizeProperty mOffset1, LinePropertyOffset1, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get Offset1() As Size
Const ProcName As String = "Offset1"
On Error GoTo Err

Set Offset1 = mOffset1

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Offset2(ByVal Value As Size)

Const ProcName As String = "Offset2"
On Error GoTo Err

setSizeProperty mOffset2, LinePropertyOffset2, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Get Offset2() As Size
Const ProcName As String = "Offset2"
On Error GoTo Err

Set Offset2 = mOffset2

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Thickness(ByVal Value As Long)
Const ProcName As String = "Thickness"

On Error GoTo Err

setLongProperty mThickness, LinePropertyThickness, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Thickness() As Long
Const ProcName As String = "Thickness"
On Error GoTo Err

checkPropertySet LinePropertyThickness
Thickness = mThickness

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearProperty( _
                ByVal pPropFlag As LinePropertyFlags)
Const ProcName As String = "ClearProperty"

On Error GoTo Err

mPropertyFlags = gClearFlag(mPropertyFlags, pPropFlag)
If Not mConfig Is Nothing Then mConfig.RemoveSetting propertyFlagToConfigSetting(pPropFlag)
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function clone() As LineStyle
Const ProcName As String = "clone"

On Error GoTo Err

Set clone = New LineStyle
If IsPropertySet(LinePropertyArrowEndColor) Then clone.ArrowEndColor = mArrowEndColor
If IsPropertySet(LinePropertyArrowEndFillColor) Then clone.ArrowEndFillColor = mArrowEndFillColor
If IsPropertySet(LinePropertyArrowEndFillStyle) Then clone.ArrowEndFillStyle = mArrowEndFillStyle
If IsPropertySet(LinePropertyArrowEndLength) Then clone.ArrowEndLength = mArrowEndLength
If IsPropertySet(LinePropertyArrowEndStyle) Then clone.ArrowEndStyle = mArrowEndStyle
If IsPropertySet(LinePropertyArrowEndWidth) Then clone.ArrowEndWidth = mArrowEndWidth
If IsPropertySet(LinePropertyArrowStartColor) Then clone.ArrowStartColor = mArrowStartColor
If IsPropertySet(LinePropertyArrowStartFillColor) Then clone.ArrowStartFillColor = mArrowStartFillColor
If IsPropertySet(LinePropertyArrowStartFillStyle) Then clone.ArrowStartFillStyle = mArrowStartfillStyle
If IsPropertySet(LinePropertyArrowStartLength) Then clone.ArrowStartLength = mArrowStartLength
If IsPropertySet(LinePropertyArrowStartStyle) Then clone.ArrowStartStyle = mArrowStartStyle
If IsPropertySet(LinePropertyArrowStartWidth) Then clone.ArrowStartWidth = mArrowStartWidth
If IsPropertySet(LinePropertyColor) Then clone.Color = mColor
If IsPropertySet(LinePropertyExtendAfter) Then clone.ExtendAfter = mExtendAfter
If IsPropertySet(LinePropertyExtendBefore) Then clone.ExtendBefore = mExtendBefore
If IsPropertySet(LinePropertyExtended) Then clone.Extended = mExtended
If IsPropertySet(LinePropertyFixedX) Then clone.FixedX = mFixedX
If IsPropertySet(LinePropertyFixedY) Then clone.FixedY = mFixedY
If IsPropertySet(LinePropertyIncludeInAutoscale) Then clone.IncludeInAutoscale = mIncludeInAutoscale
If IsPropertySet(LinePropertyLayer) Then clone.Layer = mLayer
If IsPropertySet(LinePropertyLineStyle) Then clone.LineStyle = mLineStyle
If IsPropertySet(LinePropertyOffset1) Then clone.Offset1 = mOffset1
If IsPropertySet(LinePropertyOffset2) Then clone.Offset2 = mOffset2
If IsPropertySet(LinePropertyThickness) Then clone.Thickness = mThickness

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function IsPropertySet( _
                ByVal pPropertyFlag As LinePropertyFlags) As Boolean
                
IsPropertySet = gIsFlagSet(mPropertyFlags, pPropertyFlag)

End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = pConfig
mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)

If mConfig.GetSetting(ConfigSettingArrowEndColor) <> "" Then ArrowEndColor = mConfig.GetSetting(ConfigSettingArrowEndColor)
If mConfig.GetSetting(ConfigSettingArrowEndFillColor) <> "" Then ArrowEndFillColor = mConfig.GetSetting(ConfigSettingArrowEndFillColor)
If mConfig.GetSetting(ConfigSettingArrowEndFillStyle) <> "" Then ArrowEndFillStyle = mConfig.GetSetting(ConfigSettingArrowEndFillStyle)
If mConfig.GetSetting(ConfigSettingArrowEndLength) <> "" Then ArrowEndLength = mConfig.GetSetting(ConfigSettingArrowEndLength)
If mConfig.GetSetting(ConfigSettingArrowEndStyle) <> "" Then ArrowEndStyle = mConfig.GetSetting(ConfigSettingArrowEndStyle)
If mConfig.GetSetting(ConfigSettingArrowEndWidth) <> "" Then ArrowEndWidth = mConfig.GetSetting(ConfigSettingArrowEndWidth)

If mConfig.GetSetting(ConfigSettingArrowStartColor) <> "" Then ArrowStartColor = mConfig.GetSetting(ConfigSettingArrowStartColor)
If mConfig.GetSetting(ConfigSettingArrowStartFillColor) <> "" Then ArrowStartFillColor = mConfig.GetSetting(ConfigSettingArrowStartFillColor)
If mConfig.GetSetting(ConfigSettingArrowStartfillStyle) <> "" Then ArrowStartFillStyle = mConfig.GetSetting(ConfigSettingArrowStartfillStyle)
If mConfig.GetSetting(ConfigSettingArrowStartLength) <> "" Then ArrowStartLength = mConfig.GetSetting(ConfigSettingArrowStartLength)
If mConfig.GetSetting(ConfigSettingArrowStartStyle) <> "" Then ArrowStartStyle = mConfig.GetSetting(ConfigSettingArrowStartStyle)
If mConfig.GetSetting(ConfigSettingArrowStartWidth) <> "" Then ArrowStartWidth = mConfig.GetSetting(ConfigSettingArrowStartWidth)

If mConfig.GetSetting(ConfigSettingColor) <> "" Then Color = mConfig.GetSetting(ConfigSettingColor)

If mConfig.GetSetting(ConfigSettingExtendAfter) <> "" Then ExtendAfter = mConfig.GetSetting(ConfigSettingExtendAfter)
If mConfig.GetSetting(ConfigSettingExtendBefore) <> "" Then ExtendBefore = mConfig.GetSetting(ConfigSettingExtendBefore)
If mConfig.GetSetting(ConfigSettingExtended) <> "" Then Extended = mConfig.GetSetting(ConfigSettingExtended)

If mConfig.GetSetting(ConfigSettingFixedX) <> "" Then FixedX = mConfig.GetSetting(ConfigSettingFixedX)
If mConfig.GetSetting(ConfigSettingFixedY) <> "" Then FixedY = mConfig.GetSetting(ConfigSettingFixedY)

If mConfig.GetSetting(ConfigSettingIncludeInAutoscale) <> "" Then IncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)

If mConfig.GetSetting(ConfigSettingLayer) <> "" Then Layer = mConfig.GetSetting(ConfigSettingLayer)
If mConfig.GetSetting(ConfigSettingLineStyle) <> "" Then LineStyle = mConfig.GetSetting(ConfigSettingLineStyle)

If mConfig.GetSetting(ConfigSettingThickness) <> "" Then Thickness = mConfig.GetSetting(ConfigSettingThickness)

If Not mConfig.GetConfigurationSection(ConfigSectionOffset1) Is Nothing Then Offset1 = gLoadSizeFromConfig(mConfig.GetConfigurationSection(ConfigSectionOffset1))
If Not mConfig.GetConfigurationSection(ConfigSectionOffset2) Is Nothing Then Offset2 = gLoadSizeFromConfig(mConfig.GetConfigurationSection(ConfigSectionOffset2))

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub checkPropertySet( _
                ByVal pPropFlag As LinePropertyFlags)
Const ProcName As String = "checkPropertySet"
On Error GoTo Err

If Not IsPropertySet(pPropFlag) Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Property is not set"
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEventData

Const ProcName As String = "firePropertyChanged"

On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function propertyFlagToConfigSetting( _
            ByVal pPropFlag As LinePropertyFlags) As String
Const ProcName As String = "propertyFlagToConfigSetting"
On Error GoTo Err

Select Case pPropFlag
Case LinePropertyColor
    propertyFlagToConfigSetting = ConfigSettingColor
Case LinePropertyThickness
    propertyFlagToConfigSetting = ConfigSettingThickness
Case LinePropertyLineStyle
    propertyFlagToConfigSetting = ConfigSettingLineStyle
Case LinePropertyExtendBefore
    propertyFlagToConfigSetting = ConfigSettingExtendBefore
Case LinePropertyExtendAfter
    propertyFlagToConfigSetting = ConfigSettingExtendAfter
Case LinePropertyArrowStartStyle
    propertyFlagToConfigSetting = ConfigSettingArrowStartStyle
Case LinePropertyArrowStartLength
    propertyFlagToConfigSetting = ConfigSettingArrowStartLength
Case LinePropertyArrowStartWidth
    propertyFlagToConfigSetting = ConfigSettingArrowStartWidth
Case LinePropertyArrowStartColor
    propertyFlagToConfigSetting = ConfigSettingArrowStartColor
Case LinePropertyArrowStartFillColor
    propertyFlagToConfigSetting = ConfigSettingArrowStartFillColor
Case LinePropertyArrowStartFillStyle
    propertyFlagToConfigSetting = ConfigSettingArrowStartfillStyle
Case LinePropertyArrowEndStyle
    propertyFlagToConfigSetting = ConfigSettingArrowEndStyle
Case LinePropertyArrowEndLength
    propertyFlagToConfigSetting = ConfigSettingArrowEndLength
Case LinePropertyArrowEndWidth
    propertyFlagToConfigSetting = ConfigSettingArrowEndWidth
Case LinePropertyArrowEndColor
    propertyFlagToConfigSetting = ConfigSettingArrowEndColor
Case LinePropertyArrowEndFillColor
    propertyFlagToConfigSetting = ConfigSettingArrowEndFillColor
Case LinePropertyArrowEndFillStyle
    propertyFlagToConfigSetting = ConfigSettingArrowEndFillStyle
Case LinePropertyFixedX
    propertyFlagToConfigSetting = ConfigSettingFixedX
Case LinePropertyFixedY
    propertyFlagToConfigSetting = ConfigSettingFixedY
Case LinePropertyIncludeInAutoscale
    propertyFlagToConfigSetting = ConfigSettingIncludeInAutoscale
Case LinePropertyExtended
    propertyFlagToConfigSetting = ConfigSettingExtended
Case LinePropertyLayer
    propertyFlagToConfigSetting = ConfigSettingLayer
Case LinePropertyOffset1
    propertyFlagToConfigSetting = ConfigSectionOffset1
Case LinePropertyOffset2
    propertyFlagToConfigSetting = ConfigSectionOffset2
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid property flag value"
End Select

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSizeProperty( _
                ByRef pProp As Size, _
                ByVal pPropFlag As LinePropertyFlags, _
                ByVal Value As Size)
Const ProcName As String = "setSizeProperty"
On Error GoTo Err

If pProp Is Value And IsPropertySet(pPropFlag) Then Exit Sub
If Not pProp Is Nothing Then pProp.RemoveFromConfig
Set pProp = Value
If pProp Is Nothing Then
    ClearProperty pPropFlag
Else
    setPropertyFlag pPropFlag
    If Not mConfig Is Nothing Then pProp.ConfigurationSection = mConfig.AddConfigurationSection(propertyFlagToConfigSetting(pPropFlag))
End If
firePropertyChanged gLinePropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As LinePropertyFlags)
Const ProcName As String = "setPropertyFlag"

On Error GoTo Err

mPropertyFlags = gSetFlag(mPropertyFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



