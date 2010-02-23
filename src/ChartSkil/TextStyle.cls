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

'@================================================================================
' Events
'@================================================================================

Event PropertyChanged( _
                ByRef ev As PropertyChangedEvent)
                
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

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

Set mFont = New StdFont
mFont.Bold = False
mFont.Italic = False
mFont.Name = "Arial"
mFont.Size = 8
mFont.Strikethrough = False
mFont.Underline = False

mColor = vbBlack
mBox = False
mBoxColor = vbBlack
mBoxStyle = LineStyles.LineSolid
mBoxThickness = 1
mBoxFillColor = vbWhite
mBoxFillStyle = FillStyles.FillSolid
mBoxFillWithBackgroundColor = False
mAlign = TextAlignModes.AlignTopLeft
mPaddingX = 1
'mPaddingY = 0.5
mIncludeInAutoscale = False
mFixedX = False
mFixedY = False
mExtended = False

mAngle = 0
mJustification = TextJustifyModes.JustifyLeft
mMultiLine = False
mEllipsis = EllipsisModes.EllipsisNone
mExpandTabs = True
mTabWidth = 8
mWordWrap = True
mHideIfBlank = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFont Event Handlers
'@================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
firePropertyChanged "Font"
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Align(ByVal value As TextAlignModes)
Const ProcName As String = "Align"
Dim failpoint As String
On Error GoTo Err

If mAlign = value Then Exit Property
mAlign = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAlign, mAlign
firePropertyChanged "Align"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Align() As TextAlignModes
Align = mAlign
End Property

Public Property Let Angle(ByVal value As Double)
Const ProcName As String = "Angle"
Dim failpoint As String
On Error GoTo Err

If mAngle = value Then Exit Property
mAngle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAngle, mAngle
firePropertyChanged "Angle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Angle() As Double
Angle = mAngle
End Property

Public Property Let Box(ByVal value As Boolean)
Const ProcName As String = "Box"
Dim failpoint As String
On Error GoTo Err

If mBox = value Then Exit Property
mBox = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBox, mBox
firePropertyChanged "Box"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Box() As Boolean
Box = mBox
End Property

Public Property Let BoxColor(ByVal value As Long)
Const ProcName As String = "BoxColor"
Dim failpoint As String
On Error GoTo Err

If mBoxColor = value Then Exit Property
mBoxColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxColor, mBoxColor
firePropertyChanged "BoxColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxColor() As Long
BoxColor = mBoxColor
End Property

Public Property Let BoxFillColor(ByVal value As Long)
Const ProcName As String = "BoxFillColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillColor = value Then Exit Property
mBoxFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillColor, mBoxFillColor
firePropertyChanged "BoxFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillColor() As Long
BoxFillColor = mBoxFillColor
End Property

Public Property Let BoxFillStyle(ByVal value As FillStyles)
Const ProcName As String = "BoxFillStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxFillStyle = value Then Exit Property
mBoxFillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillStyle, mBoxFillStyle
firePropertyChanged "BoxFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillStyle() As FillStyles
BoxFillStyle = mBoxFillStyle
End Property

Public Property Let BoxFillWithBackgroundColor(ByVal value As Boolean)
Const ProcName As String = "BoxFillWithBackgroundColor"
Dim failpoint As String
On Error GoTo Err

If mBoxFillWithBackgroundColor = value Then Exit Property
mBoxFillWithBackgroundColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxFillWithBackgroundColor, mBoxFillWithBackgroundColor
firePropertyChanged "BoxFillWithBackgroundColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxFillWithBackgroundColor() As Boolean
BoxFillWithBackgroundColor = mBoxFillWithBackgroundColor
End Property

Public Property Let BoxStyle(ByVal value As LineStyles)
Const ProcName As String = "BoxStyle"
Dim failpoint As String
On Error GoTo Err

If mBoxStyle = value Then Exit Property
mBoxStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxStyle, mBoxStyle
firePropertyChanged "BoxStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxStyle() As LineStyles
BoxStyle = mBoxStyle
End Property

Public Property Let BoxThickness(ByVal value As Long)
Const ProcName As String = "BoxThickness"
Dim failpoint As String
On Error GoTo Err

If mBoxThickness = value Then Exit Property
mBoxThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBoxThickness, mBoxThickness
firePropertyChanged "BoxThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get BoxThickness() As Long
BoxThickness = mBoxThickness
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
If mColor = value Then Exit Property
mColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingColor, mColor
firePropertyChanged "Color"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Color() As Long
Color = mColor
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

Set mConfig = value

SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
mConfig.SetSetting ConfigSettingAlign, mAlign
mConfig.SetSetting ConfigSettingAngle, mAngle
mConfig.SetSetting ConfigSettingBox, mBox
mConfig.SetSetting ConfigSettingBoxColor, mBoxColor
mConfig.SetSetting ConfigSettingBoxFillColor, mBoxFillColor
mConfig.SetSetting ConfigSettingBoxFillStyle, mBoxFillStyle
mConfig.SetSetting ConfigSettingBoxFillWithBackgroundColor, mBoxFillWithBackgroundColor
mConfig.SetSetting ConfigSettingBoxStyle, mBoxStyle
mConfig.SetSetting ConfigSettingBoxThickness, mBoxThickness
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingExtended, mExtended
mConfig.SetSetting ConfigSettingFixedX, mFixedX
mConfig.SetSetting ConfigSettingFixedY, mFixedY
mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
mConfig.SetSetting ConfigSettingPaddingX, mPaddingX
mConfig.SetSetting ConfigSettingPaddingY, mPaddingY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Let Ellipsis(ByVal value As EllipsisModes)
Const ProcName As String = "Ellipsis"
Dim failpoint As String
On Error GoTo Err

If mEllipsis = value Then Exit Property
mEllipsis = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingEllipsis, mEllipsis
firePropertyChanged "Ellipsis"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Ellipsis() As EllipsisModes
Ellipsis = mEllipsis
End Property

Public Property Let ExpandTabs(ByVal value As Boolean)
Const ProcName As String = "ExpandTabs"
Dim failpoint As String
On Error GoTo Err

If mExpandTabs = value Then Exit Property
mExpandTabs = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExpandTabs, mExpandTabs
firePropertyChanged "ExpandTabs"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExpandTabs() As Boolean
ExpandTabs = mExpandTabs
End Property

Public Property Let Extended(ByVal value As Boolean)
Const ProcName As String = "Extended"
Dim failpoint As String
On Error GoTo Err

If mExtended = value Then Exit Property
mExtended = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtended, mExtended
firePropertyChanged "Extended"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Extended() As Boolean
Extended = mExtended
End Property

Public Property Let FixedX(ByVal value As Boolean)
Const ProcName As String = "FixedX"
Dim failpoint As String
On Error GoTo Err

If mFixedX = value Then Exit Property
mFixedX = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingFixedX, mFixedX
firePropertyChanged "FixedX"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedX() As Boolean
FixedX = mFixedX
End Property

Public Property Let FixedY(ByVal value As Boolean)
Const ProcName As String = "FixedY"
Dim failpoint As String
On Error GoTo Err

If mFixedY = value Then Exit Property
mFixedY = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingFixedY, mFixedY
firePropertyChanged "FixedY"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get FixedY() As Boolean
FixedY = mFixedY
End Property

Public Property Let Font(ByVal value As StdFont)
Const ProcName As String = "Font"
Dim failpoint As String
On Error GoTo Err

If mFont Is value Then Exit Property
Set mFont = value
If Not mConfig Is Nothing Then SaveFontToConfig mFont, mConfig.AddConfigurationSection(ConfigSectionFont)
firePropertyChanged "Font"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Font() As StdFont
Set Font = mFont
End Property

Public Property Let HideIfBlank(ByVal value As Boolean)
Const ProcName As String = "HideIfBlank"
Dim failpoint As String
On Error GoTo Err

If mHideIfBlank = value Then Exit Property
mHideIfBlank = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHideIfBlank, mHideIfBlank
firePropertyChanged "HideIfBlank"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HideIfBlank() As Boolean
HideIfBlank = mHideIfBlank
End Property

Public Property Let IncludeInAutoscale(ByVal value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
Dim failpoint As String
On Error GoTo Err

If mIncludeInAutoscale = value Then Exit Property
mIncludeInAutoscale = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
firePropertyChanged "IncludeInAutoscale"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get IncludeInAutoscale() As Boolean
IncludeInAutoscale = mIncludeInAutoscale
End Property

Public Property Let Justification(ByVal value As TextJustifyModes)
Const ProcName As String = "Justification"
Dim failpoint As String
On Error GoTo Err

If mJustification = value Then Exit Property
mJustification = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingJustification, mJustification
firePropertyChanged "Justification"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Justification() As TextJustifyModes
Justification = mJustification
End Property

Public Property Let MultiLine(ByVal value As Boolean)
Const ProcName As String = "MultiLine"
Dim failpoint As String
On Error GoTo Err

If mMultiLine = value Then Exit Property
mMultiLine = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingMultiline, mMultiLine
firePropertyChanged "MultiLine"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get MultiLine() As Boolean
MultiLine = mMultiLine
End Property

Public Property Let PaddingX(ByVal value As Double)
Const ProcName As String = "PaddingX"
Dim failpoint As String
On Error GoTo Err

If mPaddingX = value Then Exit Property
mPaddingX = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPaddingX, mPaddingX
firePropertyChanged "PaddingX"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingX() As Double
PaddingX = mPaddingX
End Property

Public Property Let PaddingY(ByVal value As Double)
Const ProcName As String = "PaddingY"
Dim failpoint As String
On Error GoTo Err

If mPaddingY = value Then Exit Property
mPaddingY = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPaddingY, mPaddingY
firePropertyChanged "PaddingY"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PaddingY() As Double
PaddingY = mPaddingY
End Property

Public Property Let TabWidth(ByVal value As Long)
Const ProcName As String = "TabWidth"
Dim failpoint As String
On Error GoTo Err

If mTabWidth = value Then Exit Property
mTabWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingTabWidth, mTabWidth
firePropertyChanged "TabWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get TabWidth() As Long
TabWidth = mTabWidth
End Property

Public Property Let WordWrap(ByVal value As Boolean)
Const ProcName As String = "WordWrap"
Dim failpoint As String
On Error GoTo Err

If mWordWrap = value Then Exit Property
mWordWrap = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingWordWrap, mWordWrap
firePropertyChanged "WordWrap"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get WordWrap() As Boolean
WordWrap = mWordWrap
End Property

'@================================================================================
' Methods
'@================================================================================

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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Function

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = config

Set mFont = LoadFontFromConfig(mConfig.AddConfigurationSection(ConfigSectionFont))

mAlign = mConfig.GetSetting(ConfigSettingAlign, CStr(AlignTopLeft))
mAngle = mConfig.GetSetting(ConfigSettingAngle, "0")
mColor = mConfig.GetSetting(ConfigSettingColor, "0")
mBox = mConfig.GetSetting(ConfigSettingBox, "False")
mBoxColor = mConfig.GetSetting(ConfigSettingBoxColor, "0")
mBoxFillColor = mConfig.GetSetting(ConfigSettingBoxFillColor, CStr(vbWhite))
mBoxFillStyle = mConfig.GetSetting(ConfigSettingBoxFillStyle, CStr(FillStyles.FillSolid))
mBoxFillWithBackgroundColor = mConfig.GetSetting(ConfigSettingBoxFillWithBackgroundColor, "False")
mBoxStyle = mConfig.GetSetting(ConfigSettingBoxStyle, CStr(LineStyles.LineSolid))
mBoxThickness = mConfig.GetSetting(ConfigSettingBoxThickness, "1")
mExtended = mConfig.GetSetting(ConfigSettingExtended, "False")
mFixedX = mConfig.GetSetting(ConfigSettingFixedX, "False")
mFixedY = mConfig.GetSetting(ConfigSettingFixedY, "False")
mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale, "False")
mPaddingX = mConfig.GetSetting(ConfigSettingPaddingX, "1")
mPaddingY = mConfig.GetSetting(ConfigSettingPaddingY, "0")
mJustification = mConfig.GetSetting(ConfigSettingJustification, CStr(TextJustifyModes.JustifyLeft))
mMultiLine = mConfig.GetSetting(ConfigSettingMultiline, "False")
mEllipsis = mConfig.GetSetting(ConfigSettingEllipsis, CStr(EllipsisModes.EllipsisNone))
mExpandTabs = mConfig.GetSetting(ConfigSettingExpandTabs, "True")
mTabWidth = mConfig.GetSetting(ConfigSettingTabWidth, "8")
mWordWrap = mConfig.GetSetting(ConfigSettingWordWrap, "True")
mHideIfBlank = mConfig.GetSetting(ConfigSettingHideIfBlank, "True")
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

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub firePropertyChanged( _
                ByVal Name As String)
Dim ev As PropertyChangedEvent

Const ProcName As String = "firePropertyChanged"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.PropertyName = Name
RaiseEvent PropertyChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

