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

Private Const ModuleName                As String = "LineStyle"

Private Const ConfigSettingColor As String = "&Color"
Private Const ConfigSettingThickness As String = "&Thickness"
Private Const ConfigSettingLineStyle  As String = "&LineStyle"
Private Const ConfigSettingExtendBefore As String = "&ExtendBefore"
Private Const ConfigSettingExtendAfter As String = "&ExtendAfter"
Private Const ConfigSettingArrowStartStyle  As String = "&ArrowStartStyle"
Private Const ConfigSettingArrowStartLength As String = "&ArrowStartLength"
Private Const ConfigSettingArrowStartWidth As String = "&ArrowStartWidth"
Private Const ConfigSettingArrowStartColor As String = "&ArrowStartColor"
Private Const ConfigSettingArrowStartFillColor As String = "&ArrowStartFillColor"
Private Const ConfigSettingArrowStartfillStyle  As String = "&ArrowStartFillStyle"
Private Const ConfigSettingArrowEndStyle  As String = "&ArrowEndStyle"
Private Const ConfigSettingArrowEndLength As String = "&ArrowEndLength"
Private Const ConfigSettingArrowEndWidth As String = "&ArrowEndWidth"
Private Const ConfigSettingArrowEndColor As String = "&ArrowEndColor"
Private Const ConfigSettingArrowEndFillColor As String = "&ArrowEndFillColor"
Private Const ConfigSettingArrowEndFillStyle  As String = "&ArrowEndFillStyle"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"
Private Const ConfigSettingFixedX               As String = "&FixedX"
Private Const ConfigSettingFixedY               As String = "&FixedY"
Private Const ConfigSettingExtended             As String = "&Extended"

'@================================================================================
' Member variables
'@================================================================================

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

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mColor = vbBlack
mThickness = 1
mLineStyle = LineStyles.LineSolid
mExtendBefore = False
mExtendAfter = False
mArrowStartStyle = ArrowStyles.ArrowNone
mArrowStartLength = 10
mArrowStartWidth = 10
mArrowStartColor = vbBlack
mArrowStartFillColor = vbBlack
mArrowStartfillStyle = FillStyles.FillSolid
mArrowEndStyle = ArrowStyles.ArrowNone
mArrowEndLength = 10
mArrowEndWidth = 10
mArrowEndColor = vbBlack
mArrowEndFillColor = vbBlack
mArrowEndFillStyle = FillStyles.FillSolid
mIncludeInAutoscale = False
mFixedX = False
mFixedY = False
mExtended = False
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let ArrowStartColor(ByVal value As Long)
Const ProcName As String = "ArrowStartColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
mArrowStartColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartColor, mArrowStartColor
firePropertyChanged "ArrowStartColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartColor() As Long
ArrowStartColor = mArrowStartColor
End Property

Public Property Let ArrowStartFillColor(ByVal value As Long)
Const ProcName As String = "ArrowStartFillColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
mArrowStartFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartFillColor, mArrowStartFillColor
firePropertyChanged "ArrowStartFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillColor() As Long
ArrowStartFillColor = mArrowStartFillColor
End Property

Public Property Let ArrowStartFillStyle(ByVal value As FillStyles)
Const ProcName As String = "ArrowStartFillStyle"
Dim failpoint As String
On Error GoTo Err

mArrowStartfillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartfillStyle, mArrowStartfillStyle
firePropertyChanged "ArrowStartFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartFillStyle() As FillStyles
ArrowStartFillStyle = mArrowStartfillStyle
End Property

Public Property Let ArrowStartLength(ByVal value As Long)
Const ProcName As String = "ArrowStartLength"
Dim failpoint As String
On Error GoTo Err

mArrowStartLength = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartLength, mArrowStartLength
firePropertyChanged "ArrowStartLength"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartLength() As Long
ArrowStartLength = mArrowStartLength
End Property

Public Property Let ArrowStartStyle(ByVal value As ArrowStyles)
Const ProcName As String = "ArrowStartStyle"
Dim failpoint As String
On Error GoTo Err

mArrowStartStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartStyle, mArrowStartStyle
firePropertyChanged "ArrowStartStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartStyle() As ArrowStyles
ArrowStartStyle = mArrowStartStyle
End Property

Public Property Let ArrowStartWidth(ByVal value As Long)
Const ProcName As String = "ArrowStartWidth"
Dim failpoint As String
On Error GoTo Err

mArrowStartWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowStartWidth, mArrowStartWidth
firePropertyChanged "ArrowStartWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowStartWidth() As Long
ArrowStartWidth = mArrowStartWidth
End Property

Public Property Let ArrowEndColor(ByVal value As Long)
Const ProcName As String = "ArrowEndColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
mArrowEndColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndColor, mArrowEndColor
firePropertyChanged "ArrowEndColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndColor() As Long
ArrowEndColor = mArrowEndColor
End Property

Public Property Let ArrowEndFillColor(ByVal value As Long)
Const ProcName As String = "ArrowEndFillColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
mArrowEndFillColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
firePropertyChanged "ArrowEndFillColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillColor() As Long
ArrowEndFillColor = mArrowEndFillColor
End Property

Public Property Let ArrowEndFillStyle(ByVal value As FillStyles)
Const ProcName As String = "ArrowEndFillStyle"
Dim failpoint As String
On Error GoTo Err

mArrowEndFillStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
firePropertyChanged "ArrowEndFillStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndFillStyle() As FillStyles
ArrowEndFillStyle = mArrowEndFillStyle
End Property

Public Property Let ArrowEndLength(ByVal value As Long)
Const ProcName As String = "ArrowEndLength"
Dim failpoint As String
On Error GoTo Err

mArrowEndLength = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndLength, mArrowEndLength
firePropertyChanged "ArrowEndLength"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndLength() As Long
ArrowEndLength = mArrowEndLength
End Property

Public Property Let ArrowEndStyle(ByVal value As ArrowStyles)
Const ProcName As String = "ArrowEndStyle"
Dim failpoint As String
On Error GoTo Err

mArrowEndStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndStyle, mArrowEndStyle
firePropertyChanged "ArrowEndStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndStyle() As ArrowStyles
ArrowEndStyle = mArrowEndStyle
End Property

Public Property Let ArrowEndWidth(ByVal value As Long)
Const ProcName As String = "ArrowEndWidth"
Dim failpoint As String
On Error GoTo Err

mArrowEndWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingArrowEndWidth, mArrowEndWidth
firePropertyChanged "ArrowEndWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ArrowEndWidth() As Long
ArrowEndWidth = mArrowEndWidth
End Property

Public Property Let Color(ByVal value As Long)
Const ProcName As String = "Color"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color"
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

mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingThickness, mThickness
mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
mConfig.SetSetting ConfigSettingExtendBefore, mExtendBefore
mConfig.SetSetting ConfigSettingExtendAfter, mExtendAfter
mConfig.SetSetting ConfigSettingArrowStartStyle, mArrowStartStyle
mConfig.SetSetting ConfigSettingArrowStartLength, mArrowStartLength
mConfig.SetSetting ConfigSettingArrowStartWidth, mArrowStartWidth
mConfig.SetSetting ConfigSettingArrowStartColor, mArrowStartColor
mConfig.SetSetting ConfigSettingArrowStartFillColor, mArrowStartFillColor
mConfig.SetSetting ConfigSettingArrowStartfillStyle, mArrowStartfillStyle
mConfig.SetSetting ConfigSettingArrowEndStyle, mArrowEndStyle
mConfig.SetSetting ConfigSettingArrowEndLength, mArrowEndLength
mConfig.SetSetting ConfigSettingArrowEndWidth, mArrowEndWidth
mConfig.SetSetting ConfigSettingArrowEndColor, mArrowEndColor
mConfig.SetSetting ConfigSettingArrowEndFillColor, mArrowEndFillColor
mConfig.SetSetting ConfigSettingArrowEndFillStyle, mArrowEndFillStyle
mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale
mConfig.SetSetting ConfigSettingExtended, mExtended
mConfig.SetSetting ConfigSettingFixedX, mFixedX
mConfig.SetSetting ConfigSettingFixedY, mFixedY

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Let ExtendAfter(ByVal value As Boolean)
Const ProcName As String = "ExtendAfter"
Dim failpoint As String
On Error GoTo Err

mExtendAfter = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtendAfter, mExtendAfter
firePropertyChanged "ExtendAfter"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendAfter() As Boolean
ExtendAfter = mExtendAfter
End Property

Public Property Let ExtendBefore(ByVal value As Boolean)
Const ProcName As String = "ExtendBefore"
Dim failpoint As String
On Error GoTo Err

mExtendBefore = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingExtendBefore, mExtendBefore
firePropertyChanged "ExtendBefore"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get ExtendBefore() As Boolean
ExtendBefore = mExtendBefore
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

Public Property Let LineStyle(ByVal value As LineStyles)
Const ProcName As String = "LineStyle"
Dim failpoint As String
On Error GoTo Err

mLineStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
firePropertyChanged "LineStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mLineStyle
End Property

Public Property Let Thickness(ByVal value As Long)
Const ProcName As String = "Thickness"
Dim failpoint As String
On Error GoTo Err

mThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingThickness, mThickness
firePropertyChanged "Thickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get Thickness() As Long
Thickness = mThickness
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As LineStyle
Const ProcName As String = "clone"
Dim failpoint As String
On Error GoTo Err

Set clone = New LineStyle
clone.ArrowEndColor = mArrowEndColor
clone.ArrowEndFillColor = mArrowEndFillColor
clone.ArrowEndFillStyle = mArrowEndFillStyle
clone.ArrowEndLength = mArrowEndLength
clone.ArrowEndStyle = mArrowEndStyle
clone.ArrowEndWidth = mArrowEndWidth
clone.ArrowStartColor = mArrowStartColor
clone.ArrowStartFillColor = mArrowStartFillColor
clone.ArrowStartFillStyle = mArrowStartfillStyle
clone.ArrowStartLength = mArrowStartLength
clone.ArrowStartStyle = mArrowStartStyle
clone.ArrowStartWidth = mArrowStartWidth
clone.Color = mColor
clone.ExtendAfter = mExtendAfter
clone.ExtendBefore = mExtendBefore
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.Extended = mExtended
clone.FixedX = mFixedX
clone.FixedY = mFixedY
clone.LineStyle = mLineStyle
clone.Thickness = mThickness

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
mColor = mConfig.GetSetting(ConfigSettingColor, CStr(vbBlack))
mThickness = mConfig.GetSetting(ConfigSettingThickness, "1")
mLineStyle = mConfig.GetSetting(ConfigSettingLineStyle, CStr(LineStyles.LineSolid))
mExtendBefore = mConfig.GetSetting(ConfigSettingExtendBefore, "False")
mExtendAfter = mConfig.GetSetting(ConfigSettingExtendAfter, "False")
mArrowStartStyle = mConfig.GetSetting(ConfigSettingArrowStartStyle, CStr(ArrowStyles.ArrowNone))
mArrowStartLength = mConfig.GetSetting(ConfigSettingArrowStartLength, "10")
mArrowStartWidth = mConfig.GetSetting(ConfigSettingArrowStartWidth, "10")
mArrowStartColor = mConfig.GetSetting(ConfigSettingArrowStartColor, CStr(vbBlack))
mArrowStartFillColor = mConfig.GetSetting(ConfigSettingArrowStartFillColor, CStr(vbBlack))
mArrowStartfillStyle = mConfig.GetSetting(ConfigSettingArrowStartfillStyle, CStr(FillStyles.FillSolid))
mArrowEndStyle = mConfig.GetSetting(ConfigSettingArrowEndStyle, CStr(ArrowStyles.ArrowNone))
mArrowEndLength = mConfig.GetSetting(ConfigSettingArrowEndLength, "10")
mArrowEndWidth = mConfig.GetSetting(ConfigSettingArrowEndWidth, "10")
mArrowEndColor = mConfig.GetSetting(ConfigSettingArrowEndColor, CStr(vbBlack))
mArrowEndFillColor = mConfig.GetSetting(ConfigSettingArrowEndFillColor, CStr(vbBlack))
mArrowEndFillStyle = mConfig.GetSetting(ConfigSettingArrowEndFillStyle, CStr(FillStyles.FillSolid))
mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale, "False")
mExtended = mConfig.GetSetting(ConfigSettingExtended, "False")
mFixedX = mConfig.GetSetting(ConfigSettingFixedX, "False")
mFixedY = mConfig.GetSetting(ConfigSettingFixedY, "False")

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


