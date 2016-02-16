VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl TimeframeSpecifier 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   690
   ScaleWidth      =   2295
   Begin TWControls40.TWImageCombo TimeframeUnitsCombo 
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   476
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "TimeframeSpecifier.ctx":0000
      Text            =   ""
   End
   Begin VB.TextBox TimeframeLengthText 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label LengthLabel 
      Appearance      =   0  'Flat
      Caption         =   "Length"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Label UnitsLabel 
      Appearance      =   0  'Flat
      Caption         =   "Units"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "TimeframeSpecifier"
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

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Change()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "TimeFrameSpecifier"

Private Const PropNameBackcolor                         As String = "BackColor"
Private Const PropNameEnabled                           As String = "Enabled"
Private Const PropNameForecolor                         As String = "ForeColor"
Private Const PropNameTextBackColor                     As String = "TextBackColor"
Private Const PropNameTextForeColor                     As String = "TextForeColor"

Private Const PropDfltBackColor                         As Long = vbButtonFace
Private Const PropDfltEnabled                           As Boolean = True
Private Const PropDfltForeColor                         As Long = vbButtonText
Private Const PropDfltTextBackColor                     As Long = vbWindowBackground
Private Const PropDfltTextForeColor                     As Long = vbWindowText

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultTimePeriod                              As TimePeriod

Private mValidator                                      As ITimePeriodValidator

Private mTheme                                          As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_InitProperties()
On Error Resume Next

BackColor = PropDfltBackColor
ForeColor = PropDfltForeColor
TextBackColor = UserControl.Ambient.BackColor
TextForeColor = UserControl.Ambient.ForeColor
Enabled = PropDfltEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

BackColor = PropBag.ReadProperty(PropNameBackcolor, PropDfltBackColor)
ForeColor = PropBag.ReadProperty(PropNameForecolor, PropDfltForeColor)

Enabled = PropBag.ReadProperty(PropNameEnabled, PropDfltEnabled)

TextBackColor = PropBag.ReadProperty(PropNameTextBackColor, PropDfltTextBackColor)
TextForeColor = PropBag.ReadProperty(PropNameTextForeColor, PropDfltTextForeColor)

End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

Dim controlWidth

If UserControl.Width < 1710 Then UserControl.Width = 1710
If UserControl.Height < 2 * 315 Then UserControl.Height = 2 * 315

controlWidth = UserControl.Width - LengthLabel.Width

LengthLabel.Top = 0
TimeframeLengthText.Top = 0
TimeframeLengthText.Left = LengthLabel.Width
TimeframeLengthText.Width = controlWidth

UnitsLabel.Top = UserControl.Height - TimeframeUnitsCombo.Height
TimeframeUnitsCombo.Top = UnitsLabel.Top
TimeframeUnitsCombo.Left = LengthLabel.Width
TimeframeUnitsCombo.Width = controlWidth

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty PropNameBackcolor, BackColor, PropDfltBackColor
PropBag.WriteProperty PropNameEnabled, Enabled, PropDfltEnabled
PropBag.WriteProperty PropNameForecolor, ForeColor, PropDfltForeColor
PropBag.WriteProperty PropNameTextBackColor, TextBackColor, PropDfltTextBackColor
PropBag.WriteProperty PropNameTextForeColor, TextForeColor, PropDfltTextForeColor
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TimeframeLengthText_Change()
checkChange
End Sub

Private Sub TimeframeLengthText_KeyPress(KeyAscii As Integer)
Dim l As Long

On Error GoTo Err

If KeyAscii = vbKeyBack Then Exit Sub
If KeyAscii = vbKeyTab Then Exit Sub
If KeyAscii = vbKeyLeft Then Exit Sub
If KeyAscii = vbKeyRight Then Exit Sub

If Chr(KeyAscii) < "0" Or Chr(KeyAscii) > "9" Then KeyAscii = 0: Exit Sub
l = CLng(TimeframeLengthText & Chr(KeyAscii))
Exit Sub

Err:
KeyAscii = 0
End Sub

Private Sub TimeframeUnitsCombo_Change()
checkChange
End Sub

Private Sub TimeframeUnitsCombo_Click()
checkChange
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let BackColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "backColor"
On Error GoTo Err

UserControl.BackColor = Value

PropertyChanged PropNameBackcolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColor() As OLE_COLOR
Const ProcName As String = "backColor"
On Error GoTo Err

BackColor = UserControl.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultLength() As Long
Const ProcName As String = "DefaultLength"
On Error GoTo Err

DefaultLength = TimeframeLengthText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DefaultTimePeriod( _
                ByVal Value As TimePeriod)
Const ProcName As String = "defaultTimePeriod"
On Error GoTo Err

AssertArgument Not Value Is Nothing, "Value cannot be Nothing"

setTimeframeSelection Value.Length, Value.Units

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DefaultTimePeriod() As TimePeriod
Set DefaultTimePeriod = mDefaultTimePeriod
End Property

Public Property Get Enabled() As Boolean
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled(ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
TimeframeLengthText.Enabled = Value
TimeframeUnitsCombo.Enabled = Value
PropertyChanged PropNameEnabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

UserControl.ForeColor = Value

PropertyChanged PropNameForecolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Const ProcName As String = "ForeColor"
On Error GoTo Err

ForeColor = UserControl.ForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsTimeframeValid() As Boolean
Const ProcName As String = "IsTimeframeValid"
On Error GoTo Err

If TimeframeLengthText = "" Then Exit Property

IsTimeframeValid = mValidator.IsValidTimePeriod(TimePeriod)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let TextBackColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "TextBackColor"
On Error GoTo Err

TimeframeLengthText.BackColor = Value
TimeframeUnitsCombo.BackColor = Value

PropertyChanged PropNameTextBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextBackColor() As OLE_COLOR
TextBackColor = TimeframeLengthText.BackColor
End Property

Public Property Let TextForeColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "TextForeColor"
On Error GoTo Err

TimeframeLengthText.ForeColor = Value
TimeframeUnitsCombo.ForeColor = Value

PropertyChanged PropNameTextForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextForeColor() As OLE_COLOR
TextForeColor = TimeframeLengthText.ForeColor
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

BackColor = mTheme.BackColor
ForeColor = mTheme.ForeColor
TextBackColor = mTheme.TextBackColor
TextForeColor = mTheme.TextForeColor

gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Get TimePeriod() As TimePeriod
Const ProcName As String = "TimePeriod"
On Error GoTo Err

Set mDefaultTimePeriod = GetTimePeriod(CLng(TimeframeLengthText.Text), TimePeriodUnitsFromString(TimeframeUnitsCombo.SelectedItem.Text))
Set TimePeriod = mDefaultTimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pValidator As ITimePeriodValidator)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pValidator Is Nothing, "pValidator cannot be Nothing"

Set mValidator = pValidator
setupTimeframeUnitsCombo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddItem( _
                ByVal Value As TimePeriodUnits)
Const ProcName As String = "AddItem"
On Error GoTo Err

Dim s As String

s = TimePeriodUnitsToString(Value)
If Not mValidator.IsSupportedTimePeriodUnit(Value) Then Exit Sub
TimeframeUnitsCombo.ComboItems.Add , s, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkChange()
If TimeframeLengthText.Text <> "" And Not TimeframeUnitsCombo.SelectedItem Is Nothing Then
    RaiseEvent Change
End If
End Sub

Private Function setTimeframeSelection( _
                ByVal pLength As Long, _
                ByVal pUnits As TimePeriodUnits) As Boolean
Const ProcName As String = "setTimeframeSelection"
On Error GoTo Err

If mValidator.IsValidTimePeriod(GetTimePeriod(pLength, pUnits)) Then
    TimeframeLengthText.Text = CStr(pLength)
    Set TimeframeUnitsCombo.SelectedItem = TimeframeUnitsCombo.ComboItems.Item(TimePeriodUnitsToString(pUnits))
    setTimeframeSelection = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupTimeframeUnitsCombo()
Const ProcName As String = "setupTimeframeUnitsCombo"
On Error GoTo Err

AddItem TimePeriodSecond
AddItem TimePeriodMinute
AddItem TimePeriodHour
AddItem TimePeriodDay
AddItem TimePeriodWeek
AddItem TimePeriodMonth
AddItem TimePeriodYear
AddItem TimePeriodVolume
AddItem TimePeriodTickVolume
AddItem TimePeriodTickMovement

If setTimeframeSelection(15, TimePeriodMinute) Then
ElseIf setTimeframeSelection(1, TimePeriodHour) Then
ElseIf setTimeframeSelection(1, TimePeriodDay) Then
ElseIf setTimeframeSelection(1, TimePeriodWeek) Then
ElseIf setTimeframeSelection(1, TimePeriodMonth) Then
ElseIf setTimeframeSelection(1, TimePeriodYear) Then
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
