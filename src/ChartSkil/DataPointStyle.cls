VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointStyle"
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
' Constants
'@================================================================================

Private Const ModuleName                As String = "DataPointStyle"

Private Const ConfigSettingLineThickness        As String = "&LineThickness"
Private Const ConfigSettingColor                As String = "&Color"
Private Const ConfigSettingUpColor              As String = "&UpColor"
Private Const ConfigSettingDownColor            As String = "&DownColor"
Private Const ConfigSettingLineStyle            As String = "&LineStyle"
Private Const ConfigSettingPointStyle           As String = "&PointStyle"
Private Const ConfigSettingDisplayMode          As String = "&DisplayMode"
Private Const ConfigSettingHistWidth            As String = "&HistBarWidth"
Private Const ConfigSettingIncludeInAutoscale   As String = "&IncludeInAutoscale"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mConfig As ConfigurationSection

Private mLineThickness As Long
Private mColor As Long
Private mUpColor As Long
Private mDownColor As Long
Private mLineStyle As LineStyles
Private mPointStyle As PointStyles
Private mDisplayMode As DataPointDisplayModes
Private mHistWidth As Single
Private mIncludeInAutoscale As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mLineThickness = 1
mColor = vbBlack
mLineStyle = LineStyles.LineSolid
mPointStyle = PointRound
mDisplayMode = DataPointDisplayModes.DataPointDisplayModeLine
mHistWidth = 0.6
mDownColor = -1
mUpColor = -1
mIncludeInAutoscale = True
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

mConfig.SetSetting ConfigSettingLineThickness, mLineThickness
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingUpColor, mUpColor
mConfig.SetSetting ConfigSettingDownColor, mDownColor
mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
mConfig.SetSetting ConfigSettingPointStyle, mPointStyle
mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
mConfig.SetSetting ConfigSettingHistWidth, mHistWidth
mConfig.SetSetting ConfigSettingIncludeInAutoscale, mIncludeInAutoscale

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Property

Public Property Let DisplayMode(ByVal value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
Dim failpoint As String
On Error GoTo Err

If mDisplayMode = value Then Exit Property
mDisplayMode = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDisplayMode, mDisplayMode
firePropertyChanged "DisplayMode"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
DisplayMode = mDisplayMode
End Property

Public Property Let DownColor(ByVal value As Long)
Const ProcName As String = "DownColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"
If mDownColor = value Then Exit Property
mDownColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingDownColor, mDownColor
firePropertyChanged "DownColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get DownColor() As Long
DownColor = mDownColor
End Property

Public Property Let HistogramBarWidth(ByVal value As Single)
Const ProcName As String = "HistogramBarWidth"
Dim failpoint As String
On Error GoTo Err

If value <= 0 Or value > 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "HistogramBarWidth must be greater than zero but not greater than 1"

If mHistWidth = value Then Exit Property
mHistWidth = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHistWidth, mHistWidth
firePropertyChanged "HistogramBarWidth"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get HistogramBarWidth() As Single
HistogramBarWidth = mHistWidth
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

If mLineStyle = value Then Exit Property
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

Public Property Let LineThickness(ByVal value As Long)
Const ProcName As String = "LineThickness"
Dim failpoint As String
On Error GoTo Err

If mLineThickness = value Then Exit Property
mLineThickness = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineThickness, mLineThickness
firePropertyChanged "LineThickness"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get LineThickness() As Long
LineThickness = mLineThickness
End Property

Public Property Let PointStyle(ByVal value As PointStyles)
Const ProcName As String = "PointStyle"
Dim failpoint As String
On Error GoTo Err

If mPointStyle = value Then Exit Property
mPointStyle = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingPointStyle, mPointStyle
firePropertyChanged "PointStyle"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get PointStyle() As PointStyles
PointStyle = mPointStyle
End Property

Public Property Let UpColor(ByVal value As Long)
Const ProcName As String = "UpColor"
Dim failpoint As String
On Error GoTo Err

If Not gIsValidColor(value) And value <> -1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value is not a valid Color or -1"
If mUpColor = value Then Exit Property
mUpColor = value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingUpColor, mUpColor
firePropertyChanged "UpColor"

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get UpColor() As Long
UpColor = mUpColor
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function clone() As DataPointStyle
Const ProcName As String = "clone"
Dim failpoint As String
On Error GoTo Err

Set clone = New DataPointStyle
clone.Color = mColor
clone.DisplayMode = mDisplayMode
clone.DownColor = mDownColor
clone.HistogramBarWidth = mHistWidth
clone.IncludeInAutoscale = mIncludeInAutoscale
clone.LineStyle = mLineStyle
clone.LineThickness = mLineThickness
clone.PointStyle = PointStyle
clone.UpColor = mUpColor

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
mLineThickness = mConfig.GetSetting(ConfigSettingLineThickness, "1")
mColor = mConfig.GetSetting(ConfigSettingColor, CStr(vbBlack))
mUpColor = mConfig.GetSetting(ConfigSettingUpColor, "-1")
mDownColor = mConfig.GetSetting(ConfigSettingDownColor, "-1")
mLineStyle = mConfig.GetSetting(ConfigSettingLineStyle, CStr(LineStyles.LineSolid))
mPointStyle = mConfig.GetSetting(ConfigSettingPointStyle, CStr(PointStyles.PointRound))
mDisplayMode = mConfig.GetSetting(ConfigSettingDisplayMode, CStr(DataPointDisplayModes.DataPointDisplayModeLine))
mHistWidth = mConfig.GetSetting(ConfigSettingHistWidth, "0.6")
mIncludeInAutoscale = mConfig.GetSetting(ConfigSettingIncludeInAutoscale, "True")

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





