VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextStyle"
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

Private Const ModuleName                As String = "TextStyle"

Private Const ConfigSectionFont                     As String = "Font"

Private Const ConfigSettingAngle                    As String = "&Angle"
Private Const ConfigSettingColor                    As String = "&Color"
Private Const ConfigSettingBox                      As String = "&Box"
Private Const ConfigSettingBoxColor                 As String = "&BoxColor"
Private Const ConfigSettingBoxStyle                 As String = "&BoxStyle"
Private Const ConfigSettingBoxThickness             As String = "&BoxThickness"
Private Const ConfigSettingBoxFillColor             As String = "&BoxFillColor"
Private Const ConfigSettingBoxFillStyle             As String = "&BoxFillStyle"
Private Const ConfigSettingBoxFillWithBackgroundColor   As String = "&BoxFillWithBackgroundColor"
Private Const ConfigSettingAlign                    As String = "&Align"
Private Const ConfigSettingPaddingX                 As String = "&PaddingX"
Private Const ConfigSettingPaddingY                 As String = "&PaddingY"
Private Const ConfigSettingIncludeInAutoscale       As String = "&IncludeInAutoscale"
Private Const ConfigSettingFixedX                   As String = "&FixedX"
Private Const ConfigSettingFixedY                   As String = "&FixedY"
Private Const ConfigSettingExtended                 As String = "&Extended"
Private Const ConfigSettingJustification            As String = "&Justification"
Private Const ConfigSettingMultiline                As String = "&Multiline"
Private Const ConfigSettingEllipsis                 As String = "&Ellipsis"
Private Const ConfigSettingExpandTabs               As String = "&ExpandTabs"
Private Const ConfigSettingTabWidth                 As String = "&TabWidth"
Private Const ConfigSettingWordWrap                 As String = "&WordWrap"
Private Const ConfigSettingHideIfBlank              As String = "&HideIfBlank"
Private Const ConfigSettingLayer                    As String = "&Layer"

Private Const ConfigSectionLeftMargin               As String = "LeftMargin"
Private Const ConfigSectionOffset                   As String = "Offset"
Private Const ConfigSectionRightMargin              As String = "RightMargin"
Private Const ConfigSectionSize                     As String = "Size"

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mStyleType As String
Private mConfig As ConfigurationSection

Private WithEvents mFont As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mColor As Long
Private mBox As Boolean
Private mBoxColor As Long
Private mBoxStyle As LineStyles
Private mBoxThickness As Long
Private mBoxFillColor As Long
Private mBoxFillStyle As FillStyles
Private mBoxFillWithBackgroundColor As Boolean
Private mAlign As TextAlignModes
Private mPaddingX As Double
Private mPaddingY As Double
Private mIncludeInAutoscale As Boolean
Private mFixedX As Boolean
Private mFixedY As Boolean
Private mExtended As Boolean
Private mAngle As Double
Private mJustification As TextJustifyModes
Private mMultiLine As Boolean
Private mEllipsis As EllipsisModes
Private mExpandTabs As Boolean
Private mTabWidth As Long
Private mWordWrap As Boolean
Private mHideIfBlank As Boolean
Private mLeftMargin As Dimension
Private mRightMargin As Dimension
Private mOffset As Size
Private mSize As Size
Private mLayer As Long

Private mPropertyFlags As TextPropertyFlags

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = GenerateGUIDString
mStyleType = "Text Styles"
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
' mFont Event Handlers
'@================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
Const ProcName As String = "mFont_FontChanged"
On Error GoTo Err

If Not mConfig Is Nothing Then SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
firePropertyChanged "Font"

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Align(ByVal Value As TextAlignModes)
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

setLongProperty mAlign, TextPropertyAlign, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Align() As TextAlignModes
Const ProcName As String = "Align"
On Error GoTo Err

checkPropertySet TextPropertyAlign
Align = mAlign

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Angle(ByVal Value As Double)
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

setDoubleProperty mAngle, TextPropertyAngle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Angle() As Double
Const ProcName As String = "Angle"
On Error GoTo Err

checkPropertySet TextPropertyAngle
Angle = mAngle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Box(ByVal Value As Boolean)
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mBox, TextPropertyBox, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Box() As Boolean
Const ProcName As String = "Box"
On Error GoTo Err

checkPropertySet TextPropertyBox
Box = mBox

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxColor(ByVal Value As Long)
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

setLongProperty mBoxColor, TextPropertyBoxColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxColor() As Long
Const ProcName As String = "BoxColor"
On Error GoTo Err

checkPropertySet TextPropertyBoxColor
BoxColor = mBoxColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillColor(ByVal Value As Long)
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

setLongProperty mBoxFillColor, TextPropertyBoxFillColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillColor() As Long
Const ProcName As String = "BoxFillColor"
On Error GoTo Err

checkPropertySet TextPropertyBoxFillColor
BoxFillColor = mBoxFillColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillStyle(ByVal Value As FillStyles)
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mBoxFillStyle, TextPropertyBoxFillStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillStyle() As FillStyles
Const ProcName As String = "BoxFillStyle"
On Error GoTo Err

checkPropertySet TextPropertyBoxFillStyle
BoxFillStyle = mBoxFillStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal Value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mBoxFillWithBackgroundColor, TextPropertyBoxFillWithBackgroundColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
Const ProcName As String = "BoxFillWithBackgroundColor"
On Error GoTo Err

checkPropertySet TextPropertyBoxFillWithBackgroundColor
BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxStyle(ByVal Value As LineStyles)
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

setLongProperty mBoxStyle, TextPropertyBoxStyle, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxStyle() As LineStyles
Const ProcName As String = "BoxStyle"
On Error GoTo Err

checkPropertySet TextPropertyBoxStyle
BoxStyle = mBoxStyle

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let BoxThickness(ByVal Value As Long)
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

setLongProperty mBoxThickness, TextPropertyBoxThickness, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get BoxThickness() As Long
Const ProcName As String = "BoxThickness"
On Error GoTo Err

checkPropertySet TextPropertyBoxThickness
BoxThickness = mBoxThickness

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(Value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"

setLongProperty mColor, TextPropertyColor, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Color() As Long
Const ProcName As String = "Color"
On Error GoTo Err

checkPropertySet TextPropertyColor
Color = mColor

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

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

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If Not mFont Is Nothing Then SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)

If IsPropertySet(TextPropertyAlign) Then mConfig.SetSetting ConfigSettingAlign, mAlign
If IsPropertySet(TextPropertyAngle) Then mConfig.SetSetting ConfigSettingAngle, mAngle
If IsPropertySet(TextPropertyBox) Then mConfig.SetSetting ConfigSettingBox, mBox
If IsPropertySet(TextPropertyBoxColor) Then mConfig.SetSetting ConfigSettingBoxColor, mBoxColor
If IsPropertySet(TextPropertyBoxFillColor) Then mConfig.SetSetting ConfigSettingBoxFillColor, mBoxFillColor
If IsPropertySet(TextPropertyBoxFillStyle) Then mConfig.SetSetting ConfigSettingBoxFillStyle, mBoxFillStyle
If IsPropertySet(TextPropertyBoxFillWithBackgroundColor) Then mConfig.SetSetting ConfigSettingBoxFillWithBackgroundColor, mBoxFillWithBackgroundColor
If IsPropertySet(TextPropertyBoxStyle) Then mConfig.SetSetting ConfigSettingBoxStyle, mBoxStyle
If IsPropertySet(TextPropertyBoxThickness) Then mConfig.SetSetting ConfigSettingBoxThickness, mBoxThickness
If IsPropertySet(TextPropertyColor) Then mConfig.SetSetting ConfigSettingColor, mColor
If IsPropertySet(TextPropertyEllipsis) Then mConfig.SetSetting ConfigSettingEllipsis, mEllipsis
If IsPropertySet(TextPropertyExtended) Then mConfig.SetSetting ConfigSettingExtended, mExtended
If IsPropertySet(TextPropertyFixedX) Then mConfig.SetSetting ConfigSettingFixedX, mFixedX
If IsPropertySet(TextPropertyFixedY) Then mConfig.SetSetting ConfigSettingFixedY, mFixedY
If IsPropertySet(TextPropertyHideIfBlank) Then mConfig.SetSetting ConfigSettingHideIfBlank, mHideIfBlank
If IsPropertySet(TextPropertyIncludeInAutoscale) Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
If IsPropertySet(TextPropertyJustification) Then mConfig.SetSetting ConfigSettingJustification, mJustification
If IsPropertySet(TextPropertyLayer) Then mConfig.SetSetting ConfigSettingLayer, mLayer
If IsPropertySet(TextPropertyMultiLine) Then mConfig.SetSetting ConfigSettingMultiline, mMultiLine
If IsPropertySet(TextPropertyPaddingX) Then mConfig.SetSetting ConfigSettingPaddingX, mPaddingX
If IsPropertySet(TextPropertyPaddingY) Then mConfig.SetSetting ConfigSettingPaddingY, mPaddingY
If IsPropertySet(TextPropertyTabWidth) Then mConfig.SetSetting ConfigSettingTabWidth, mTabWidth
If IsPropertySet(TextPropertyWordWrap) Then mConfig.SetSetting ConfigSettingWordWrap, mWordWrap

If Not mOffset Is Nothing Then mOffset.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionOffset)
If Not mLeftMargin Is Nothing Then mLeftMargin.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionLeftMargin)
If Not mRightMargin Is Nothing Then mRightMargin.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionRightMargin)
If Not mSize Is Nothing Then mSize.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionSize)
Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Property

Public Property Let Ellipsis(ByVal Value As EllipsisModes)
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

setLongProperty mEllipsis, TextPropertyEllipsis, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Ellipsis() As EllipsisModes
Const ProcName As String = "Ellipsis"
On Error GoTo Err

checkPropertySet TextPropertyEllipsis
Ellipsis = mEllipsis

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ExpandTabs(ByVal Value As Boolean)
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mExpandTabs, TextPropertyExpandTabs, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ExpandTabs() As Boolean
Const ProcName As String = "ExpandTabs"
On Error GoTo Err

checkPropertySet TextPropertyExpandTabs
ExpandTabs = mExpandTabs

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Extended(ByVal Value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mExtended, TextPropertyExtended, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Extended() As Boolean
Const ProcName As String = "Extended"
On Error GoTo Err

checkPropertySet TextPropertyExtended
Extended = mExtended

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedX(ByVal Value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mFixedX, TextPropertyFixedX, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedX() As Boolean
Const ProcName As String = "FixedX"
On Error GoTo Err

checkPropertySet TextPropertyFixedX
FixedX = mFixedX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let FixedY(ByVal Value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mFixedY, TextPropertyFixedY, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get FixedY() As Boolean
Const ProcName As String = "FixedY"
On Error GoTo Err

checkPropertySet TextPropertyFixedY
FixedY = mFixedY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Font(ByVal Value As StdFont)
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

If mFont Is Value Then Exit Property
Set mFont = Value
setPropertyFlag TextPropertyFont
If Not mConfig Is Nothing Then SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
firePropertyChanged "Font"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Font() As StdFont
Const ProcName As String = "Font"
On Error GoTo Err

checkPropertySet TextPropertyFont
Set Font = mFont

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HideIfBlank(ByVal Value As Boolean)
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mHideIfBlank, TextPropertyHideIfBlank, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HideIfBlank() As Boolean
Const ProcName As String = "HideIfBlank"
On Error GoTo Err

checkPropertySet TextPropertyHideIfBlank
HideIfBlank = mHideIfBlank

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mIncludeInAutoscale, TextPropertyIncludeInAutoscale, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

checkPropertySet TextPropertyIncludeInAutoscale
IncludeInAutoscale = mIncludeInAutoscale

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Justification(ByVal Value As TextJustifyModes)
Const ProcName As String = "Justification"
Dim failpoint As String
On Error GoTo Err

setLongProperty mJustification, TextPropertyJustification, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Justification() As TextJustifyModes
Const ProcName As String = "Justification"
On Error GoTo Err

checkPropertySet TextPropertyJustification
Justification = mJustification

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
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

Public Property Let LeftMargin(ByVal Value As Dimension)
Const ProcName As String = "LeftMargin"
Dim failpoint As String
On Error GoTo Err

setDimensionProperty mLeftMargin, TextPropertyLeftMargin, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LeftMargin() As Dimension
Const ProcName As String = "LeftMargin"
On Error GoTo Err

Set LeftMargin = mLeftMargin

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let MultiLine(ByVal Value As Boolean)
Const ProcName As String = "MultiLine"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mMultiLine, TextPropertyMultiLine, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MultiLine() As Boolean
Const ProcName As String = "MultiLine"
On Error GoTo Err

checkPropertySet TextPropertyMultiLine
MultiLine = mMultiLine

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

Public Property Let Offset(ByVal Value As Size)
Const ProcName As String = "Offset"
Dim failpoint As String
On Error GoTo Err

setSizeProperty mOffset, TextPropertyOffset, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Offset() As Size
Const ProcName As String = "Offset"
On Error GoTo Err

Set Offset = mOffset

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PaddingX(ByVal Value As Double)
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

setDoubleProperty mPaddingX, TextPropertyPaddingX, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PaddingX() As Double
Const ProcName As String = "PaddingX"
On Error GoTo Err

checkPropertySet TextPropertyPaddingX
PaddingX = mPaddingX

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let PaddingY(ByVal Value As Double)
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

setDoubleProperty mPaddingY, TextPropertyPaddingY, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PaddingY() As Double
Const ProcName As String = "PaddingY"
On Error GoTo Err

checkPropertySet TextPropertyPaddingY
PaddingY = mPaddingY

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let RightMargin(ByVal Value As Dimension)
Const ProcName As String = "RightMargin"
Dim failpoint As String
On Error GoTo Err

setDimensionProperty mRightMargin, TextPropertyRightMargin, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get RightMargin() As Dimension
Const ProcName As String = "RightMargin"
On Error GoTo Err

Set RightMargin = mRightMargin

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

setSizeProperty mSize, TextPropertySize, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mSize

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let TabWidth(ByVal Value As Long)
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

setLongProperty mTabWidth, TextPropertyTabWidth, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TabWidth() As Long
Const ProcName As String = "TabWidth"
On Error GoTo Err

checkPropertySet TextPropertyTabWidth
TabWidth = mTabWidth

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let WordWrap(ByVal Value As Boolean)
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

setBooleanProperty mWordWrap, TextPropertyWordWrap, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"
On Error GoTo Err

checkPropertySet TextPropertyWordWrap
WordWrap = mWordWrap

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearProperty( _
                ByVal pPropFlag As TextPropertyFlags)
Const ProcName As String = "ClearProperty"
Dim failpoint As String
On Error GoTo Err

mPropertyFlags = gClearFlag(mPropertyFlags, pPropFlag)
If Not mConfig Is Nothing Then mConfig.RemoveSetting propertyFlagToConfigSetting(pPropFlag)
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function clone() As TextStyle
Const ProcName As String = "clone"
Dim failpoint As String
On Error GoTo Err

Set clone = New TextStyle
clone.Align = mAlign
clone.Angle = mAngle
clone.Box = mBox
clone.BoxColor = mBoxColor
clone.BoxFillColor = mBoxFillColor
clone.BoxFillStyle = mBoxFillStyle
clone.BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
clone.BoxStyle = mBoxStyle
clone.BoxThickness = mBoxThickness
clone.Color = mColor
clone.Ellipsis = mEllipsis
clone.ExpandTabs = mExpandTabs
clone.Extended = mExtended
clone.FixedX = mFixedX
clone.FixedY = mFixedY
clone.Font = gCloneFont(mFont)
clone.HideIfBlank = mHideIfBlank
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.Justification = mJustification
clone.MultiLine = mMultiLine
clone.PaddingX = mPaddingX
clone.PaddingY = mPaddingY
clone.TabWidth = mTabWidth
clone.WordWrap = mWordWrap
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function

Public Function IsPropertySet( _
                ByVal pPropertyFlag As TextPropertyFlags) As Boolean

IsPropertySet = gIsFlagSet(mPropertyFlags, pPropertyFlag)

End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig

mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)

If Not mConfig.GetConfigurationSection(ConfigSectionFont) Is Nothing Then Font = LoadFontFromConfig(mConfig.GetConfigurationSection(ConfigSectionFont))

If mConfig.GetSetting(ConfigSettingAlign) <> "" Then Align = mConfig.GetSetting(ConfigSettingAlign)
If mConfig.GetSetting(ConfigSettingAngle) <> "" Then Angle = mConfig.GetSetting(ConfigSettingAngle)

If mConfig.GetSetting(ConfigSettingColor) <> "" Then Color = mConfig.GetSetting(ConfigSettingColor)

If mConfig.GetSetting(ConfigSettingBox) <> "" Then Box = mConfig.GetSetting(ConfigSettingBox, "False")
If mConfig.GetSetting(ConfigSettingBoxColor) <> "" Then BoxColor = mConfig.GetSetting(ConfigSettingBoxColor)
If mConfig.GetSetting(ConfigSettingBoxFillColor) <> "" Then BoxFillColor = mConfig.GetSetting(ConfigSettingBoxFillColor)
If mConfig.GetSetting(ConfigSettingBoxFillStyle) <> "" Then BoxFillStyle = mConfig.GetSetting(ConfigSettingBoxFillStyle)
If mConfig.GetSetting(ConfigSettingBoxFillWithBackgroundColor) <> "" Then BoxFillWithBackgroundColor = mConfig.GetSetting(ConfigSettingBoxFillWithBackgroundColor)
If mConfig.GetSetting(ConfigSettingBoxStyle) <> "" Then BoxStyle = mConfig.GetSetting(ConfigSettingBoxStyle)
If mConfig.GetSetting(ConfigSettingBoxThickness) <> "" Then BoxThickness = mConfig.GetSetting(ConfigSettingBoxThickness)

If mConfig.GetSetting(ConfigSettingEllipsis) <> "" Then Ellipsis = mConfig.GetSetting(ConfigSettingEllipsis)
If mConfig.GetSetting(ConfigSettingExtended) <> "" Then Extended = mConfig.GetSetting(ConfigSettingExtended)
If mConfig.GetSetting(ConfigSettingExpandTabs) <> "" Then ExpandTabs = mConfig.GetSetting(ConfigSettingExpandTabs)

If mConfig.GetSetting(ConfigSettingFixedX) <> "" Then FixedX = mConfig.GetSetting(ConfigSettingFixedX)
If mConfig.GetSetting(ConfigSettingFixedY) <> "" Then FixedY = mConfig.GetSetting(ConfigSettingFixedY)

If mConfig.GetSetting(ConfigSettingHideIfBlank) <> "" Then HideIfBlank = mConfig.GetSetting(ConfigSettingHideIfBlank)

If mConfig.GetSetting(ConfigSettingIncludeInAutoscale) <> "" Then IncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale)

If mConfig.GetSetting(ConfigSettingJustification) <> "" Then Justification = mConfig.GetSetting(ConfigSettingJustification)

If mConfig.GetSetting(ConfigSettingLayer) <> "" Then Layer = mConfig.GetSetting(ConfigSettingLayer)

If mConfig.GetSetting(ConfigSettingMultiline) <> "" Then MultiLine = mConfig.GetSetting(ConfigSettingMultiline)

If mConfig.GetSetting(ConfigSettingPaddingX) <> "" Then PaddingX = mConfig.GetSetting(ConfigSettingPaddingX)
If mConfig.GetSetting(ConfigSettingPaddingY) <> "" Then PaddingY = mConfig.GetSetting(ConfigSettingPaddingY)

If mConfig.GetSetting(ConfigSettingTabWidth) <> "" Then TabWidth = mConfig.GetSetting(ConfigSettingTabWidth)

If mConfig.GetSetting(ConfigSettingWordWrap) <> "" Then WordWrap = mConfig.GetSetting(ConfigSettingWordWrap)

If Not mConfig.GetConfigurationSection(ConfigSectionLeftMargin) Is Nothing Then LeftMargin = gLoadDimensionFromConfig(mConfig.GetConfigurationSection(ConfigSectionLeftMargin))
If Not mConfig.GetConfigurationSection(ConfigSectionOffset) Is Nothing Then Offset = gLoadSizeFromConfig(mConfig.GetConfigurationSection(ConfigSectionOffset))
If Not mConfig.GetConfigurationSection(ConfigSectionRightMargin) Is Nothing Then RightMargin = gLoadDimensionFromConfig(mConfig.GetConfigurationSection(ConfigSectionRightMargin))
If Not mConfig.GetConfigurationSection(ConfigSectionSize) Is Nothing Then Size = gLoadSizeFromConfig(mConfig.GetConfigurationSection(ConfigSectionSize))

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
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
                ByVal pPropFlag As TextPropertyFlags)
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
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function propertyFlagToConfigSetting( _
            ByVal pPropFlag As TextPropertyFlags) As String
Select Case pPropFlag
Case TextPropertyColor
    propertyFlagToConfigSetting = ConfigSettingColor
Case TextPropertyBox
    propertyFlagToConfigSetting = ConfigSettingBox
Case TextPropertyBoxColor
    propertyFlagToConfigSetting = ConfigSettingBoxColor
Case TextPropertyBoxStyle
    propertyFlagToConfigSetting = ConfigSettingBoxStyle
Case TextPropertyBoxThickness
    propertyFlagToConfigSetting = ConfigSettingBoxThickness
Case TextPropertyBoxFillColor
    propertyFlagToConfigSetting = ConfigSettingBoxFillColor
Case TextPropertyBoxFillStyle
    propertyFlagToConfigSetting = ConfigSettingBoxFillStyle
Case TextPropertyAlign
    propertyFlagToConfigSetting = ConfigSettingAlign
Case TextPropertyPaddingX
    propertyFlagToConfigSetting = ConfigSettingPaddingX
Case TextPropertyPaddingY
    propertyFlagToConfigSetting = ConfigSettingPaddingY
Case TextPropertyBoxFillWithBackgroundColor
    propertyFlagToConfigSetting = ConfigSettingBoxFillWithBackgroundColor
Case TextPropertyFixedX
    propertyFlagToConfigSetting = ConfigSettingFixedX
Case TextPropertyFixedY
    propertyFlagToConfigSetting = ConfigSettingFixedY
Case TextPropertyIncludeInAutoscale
    propertyFlagToConfigSetting = ConfigSettingIncludeInAutoscale
Case TextPropertyExtended
    propertyFlagToConfigSetting = ConfigSettingExtended
Case TextPropertyAngle
    propertyFlagToConfigSetting = ConfigSettingAngle
Case TextPropertyJustification
    propertyFlagToConfigSetting = ConfigSettingJustification
Case TextPropertyMultiLine
    propertyFlagToConfigSetting = ConfigSettingMultiline
Case TextPropertyEllipsis
    propertyFlagToConfigSetting = ConfigSettingEllipsis
Case TextPropertyExpandTabs
    propertyFlagToConfigSetting = ConfigSettingExpandTabs
Case TextPropertyTabWidth
    propertyFlagToConfigSetting = ConfigSettingTabWidth
Case TextPropertyWordWrap
    propertyFlagToConfigSetting = ConfigSettingWordWrap
Case TextPropertyHideIfBlank
    propertyFlagToConfigSetting = ConfigSettingHideIfBlank
Case TextPropertyLeftMargin
    propertyFlagToConfigSetting = ConfigSectionLeftMargin
Case TextPropertyRightMargin
    propertyFlagToConfigSetting = ConfigSectionRightMargin
Case TextPropertyOffset
    propertyFlagToConfigSetting = ConfigSectionOffset
End Select
            
End Function

Private Sub setBooleanProperty( _
                ByRef pProp As Boolean, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Boolean)
Const ProcName As String = "setBooleanProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setDimensionProperty( _
                ByRef pProp As Dimension, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Dimension)
Const ProcName As String = "setDimensionProperty"
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
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setDoubleProperty( _
                ByRef pProp As Double, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setDoubleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setLongProperty( _
                ByRef pProp As Long, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Long)
Const ProcName As String = "setLongProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSingleProperty( _
                ByRef pProp As Single, _
                ByVal pPropFlag As TextPropertyFlags, _
                ByVal Value As Single)
Const ProcName As String = "setSingleProperty"
On Error GoTo Err

If pProp = Value And IsPropertySet(pPropFlag) Then Exit Sub
pProp = Value
setPropertyFlag pPropFlag
If Not mConfig Is Nothing Then mConfig.SetSetting propertyFlagToConfigSetting(pPropFlag), pProp
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setSizeProperty( _
                ByRef pProp As Size, _
                ByVal pPropFlag As TextPropertyFlags, _
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
firePropertyChanged gTextPropertyFlagToString(pPropFlag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setPropertyFlag( _
                ByVal flag As TextPropertyFlags)
Const ProcName As String = "setPropertyFlag"
Dim failpoint As String
On Error GoTo Err

mPropertyFlags = gSetFlag(mPropertyFlags, flag)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



