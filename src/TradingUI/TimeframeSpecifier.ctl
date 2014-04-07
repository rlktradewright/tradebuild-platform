VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#10.1#0"; "TWControls40.ocx"
Begin VB.UserControl TimeframeSpecifier 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   690
   ScaleWidth      =   2295
   Begin TWControls40.TWImageCombo TimeframeUnitsCombo 
      Height          =   330
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   582
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

Private Const PropNameBackColor                         As String = "BackColor"
Private Const PropNameDefaultLength                     As String = "DefaultLength"
Private Const PropNameDefaultUnits                      As String = "DefaultUnits"
Private Const PropNameEnabled                           As String = "Enabled"
Private Const PropNameForeColor                         As String = "ForeColor"

Private Const PropDfltBackColor                         As Long = vbWindowBackground
Private Const PropDfltDefaultLength                     As Long = 5
Private Const PropDfltDefaultUnits                      As Long = TimePeriodUnits.TimePeriodMinute
Private Const PropDfltEnabled                           As Boolean = True
Private Const PropDfltForeColor                         As Long = vbWindowText

'@================================================================================
' Member variables
'@================================================================================

Private mDefaultTimePeriod                              As TimePeriod

Private mValidator                                      As ITimePeriodValidator

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_InitProperties()
On Error Resume Next

UserControl.BackColor = UserControl.Ambient.BackColor
UserControl.ForeColor = UserControl.Ambient.ForeColor
Set mDefaultTimePeriod = GetTimePeriod(PropDfltDefaultLength, PropDfltDefaultUnits)
Enabled = PropDfltEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

UserControl.BackColor = UserControl.Ambient.BackColor
UserControl.ForeColor = UserControl.Ambient.ForeColor

Set DefaultTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameDefaultLength, PropDfltDefaultLength), _
                                    PropBag.ReadProperty(PropNameDefaultUnits, PropDfltDefaultUnits))
If Err.Number <> 0 Then
    DefaultTimePeriod = GetTimePeriod(PropDfltDefaultLength, PropDfltDefaultUnits)
    Err.Clear
End If

Enabled = PropBag.ReadProperty(PropNameEnabled, PropDfltEnabled)
If Err.Number <> 0 Then
    Enabled = PropDfltEnabled
    Err.Clear
End If

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
PropBag.WriteProperty PropNameBackColor, BackColor, PropDfltBackColor
PropBag.WriteProperty PropNameDefaultLength, DefaultTimePeriod.Length, PropDfltDefaultLength
PropBag.WriteProperty PropNameDefaultUnits, DefaultTimePeriod.Units, PropDfltDefaultUnits
PropBag.WriteProperty PropNameEnabled, Enabled, PropDfltEnabled
PropBag.WriteProperty PropNameForeColor, ForeColor, PropDfltForeColor
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

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
                ByVal Value As OLE_COLOR)
Const ProcName As String = "backColor"
On Error GoTo Err

TimeframeLengthText.BackColor = Value
TimeframeUnitsCombo.BackColor = Value

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
                ByVal Value As TimePeriod)
Const ProcName As String = "defaultTimePeriod"
On Error GoTo Err

AssertArgument Not Value Is Nothing, "Value cannot be Nothing"

Set mDefaultTimePeriod = Value
PropertyChanged PropNameDefaultLength
PropertyChanged PropNameDefaultUnits

TimeframeLengthText = Value.Length
setUnitsSelection Value.Units

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
Const ProcName As String = "foreColor"
On Error GoTo Err

TimeframeLengthText.ForeColor = Value
TimeframeUnitsCombo.ForeColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Const ProcName As String = "foreColor"
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

Private Sub addItem( _
                ByVal Value As TimePeriodUnits)
Const ProcName As String = "addItem"
On Error GoTo Err

Dim s As String

s = TimePeriodUnitsToString(Value)
If Not mValidator.IsSupportedTimePeriodUnit(Value) Then Exit Sub
TimeframeUnitsCombo.ComboItems.Add , s, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setUnitsSelection( _
                ByVal Value As TimePeriodUnits) As Boolean
Const ProcName As String = "setUnitsSelection"
On Error GoTo Err

If mValidator.IsValidTimePeriod(GetTimePeriod(1, Value)) Then
    TimeframeUnitsCombo.ComboItems.item(TimePeriodUnitsToString(Value)).Selected = True
    setUnitsSelection = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupTimeframeUnitsCombo()
Const ProcName As String = "setupTimeframeUnitsCombo"
On Error GoTo Err

addItem TimePeriodSecond
addItem TimePeriodMinute
addItem TimePeriodHour
addItem TimePeriodDay
addItem TimePeriodWeek
addItem TimePeriodMonth
addItem TimePeriodYear
addItem TimePeriodVolume
addItem TimePeriodTickVolume
addItem TimePeriodTickMovement
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
