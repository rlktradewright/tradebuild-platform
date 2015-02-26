VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#29.0#0"; "TWControls40.ocx"
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
      Left            =   840
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label LengthLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   855
   End
   Begin VB.Label UnitsLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
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
Private Const PropNameDefaultLength                     As String = "DefaultLength"
Private Const PropNameDefaultUnits                      As String = "DefaultUnits"
Private Const PropNameEnabled                           As String = "Enabled"
Private Const PropNameForecolor                         As String = "ForeColor"
Private Const PropNameTextBackColor                     As String = "TextBackColor"
Private Const PropNameTextForeColor                     As String = "TextForeColor"

Private Const PropDfltBackColor                         As Long = vbButtonFace
Private Const PropDfltDefaultLength                     As Long = 5
Private Const PropDfltDefaultUnits                      As Long = TimePeriodUnits.TimePeriodMinute
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
Set mDefaultTimePeriod = GetTimePeriod(PropDfltDefaultLength, PropDfltDefaultUnits)
Enabled = PropDfltEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

BackColor = PropBag.ReadProperty(PropNameBackcolor, PropDfltBackColor)
ForeColor = PropBag.ReadProperty(PropNameForecolor, PropDfltForeColor)

Set DefaultTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameDefaultLength, PropDfltDefaultLength), _
                                    PropBag.ReadProperty(PropNameDefaultUnits, PropDfltDefaultUnits))
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
PropBag.WriteProperty PropNameDefaultLength, DefaultTimePeriod.Length, PropDfltDefaultLength
PropBag.WriteProperty PropNameDefaultUnits, DefaultTimePeriod.Units, PropDfltDefaultUnits
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

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub TimeframeLengthText_Change()
RaiseEvent Change
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
RaiseEvent Change
End Sub

Private Sub TimeframeUnitsCombo_Click()
RaiseEvent Change
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let BackColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "backColor"
On Error GoTo Err

TimeframeLengthText.BackColor = value
TimeframeUnitsCombo.BackColor = value

PropertyChanged PropNameBackcolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColor() As OLE_COLOR
Const ProcName As String = "backColor"
On Error GoTo Err

BackColor = TimeframeUnitsCombo.BackColor

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
                ByVal value As TimePeriod)
Const ProcName As String = "defaultTimePeriod"
On Error GoTo Err

AssertArgument Not value Is Nothing, "Value cannot be Nothing"

Set mDefaultTimePeriod = value
PropertyChanged PropNameDefaultLength
PropertyChanged PropNameDefaultUnits

TimeframeLengthText = value.Length
setUnitsSelection value.Units

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

Public Property Let Enabled(ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = value
TimeframeLengthText.Enabled = value
TimeframeUnitsCombo.Enabled = value
PropertyChanged PropNameEnabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

TimeframeLengthText.ForeColor = value
TimeframeUnitsCombo.ForeColor = value

PropertyChanged PropNameForecolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Const ProcName As String = "ForeColor"
On Error GoTo Err

ForeColor = TimeframeUnitsCombo.ForeColor

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

Public Property Let TextBackColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "TextBackColor"
On Error GoTo Err

TimeframeLengthText.BackColor = value
TimeframeUnitsCombo.BackColor = value

PropertyChanged PropNameTextBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextBackColor() As OLE_COLOR
TextBackColor = TimeframeLengthText.BackColor
End Property

Public Property Let TextForeColor( _
                ByVal value As OLE_COLOR)
Const ProcName As String = "TextForeColor"
On Error GoTo Err

TimeframeLengthText.ForeColor = value
TimeframeUnitsCombo.ForeColor = value

PropertyChanged PropNameTextForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TextForeColor() As OLE_COLOR
TextForeColor = TimeframeLengthText.ForeColor
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

BackColor = value.BackColor
ForeColor = value.ForeColor
TextBackColor = value.TextBackColor
TextForeColor = value.TextForeColor

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

Set TimePeriod = GetTimePeriod(CLng(TimeframeLengthText.Text), TimePeriodUnitsFromString(TimeframeUnitsCombo.SelectedItem.Text))

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

DefaultTimePeriod = mDefaultTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddItem( _
                ByVal value As TimePeriodUnits)
Const ProcName As String = "AddItem"
On Error GoTo Err

Dim s As String

s = TimePeriodUnitsToString(value)
If Not mValidator.IsSupportedTimePeriodUnit(value) Then Exit Sub
TimeframeUnitsCombo.ComboItems.Add , s, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setUnitsSelection( _
                ByVal value As TimePeriodUnits) As Boolean
Const ProcName As String = "setUnitsSelection"
On Error GoTo Err

If mValidator.IsValidTimePeriod(GetTimePeriod(1, value)) Then
    Set TimeframeUnitsCombo.SelectedItem = TimeframeUnitsCombo.ComboItems.Item(TimePeriodUnitsToString(value))
    setUnitsSelection = True
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
'If setUnitsSelection(mDefaultUnits) Then
'ElseIf setUnitsSelection(TimePeriodMinute) Then
'ElseIf setUnitsSelection(TimePeriodHour) Then
'ElseIf setUnitsSelection(TimePeriodDay) Then
'ElseIf setUnitsSelection(TimePeriodWeek) Then
'ElseIf setUnitsSelection(TimePeriodMonth) Then
'Else
'    setUnitsSelection (TimePeriodYear)
'End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub
