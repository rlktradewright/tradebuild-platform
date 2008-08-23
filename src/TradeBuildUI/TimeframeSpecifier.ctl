VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#30.0#0"; "TWControls10.ocx"
Begin VB.UserControl TimeframeSpecifier 
   ClientHeight    =   690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2295
   ScaleHeight     =   690
   ScaleWidth      =   2295
   Begin TWControls10.TWImageCombo TimeframeUnitsCombo 
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

Private Const ProjectName                               As String = "TradeBuildUI25"
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

Private mDefaultTimePeriod As TimePeriod

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
setupTimeframeUnitsCombo
End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

UserControl.backColor = UserControl.Ambient.backColor
UserControl.foreColor = UserControl.Ambient.foreColor
defaultTimePeriod = GetTimePeriod(PropDfltDefaultLength, PropDfltDefaultUnits)
Enabled = PropDfltEnabled
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

UserControl.backColor = UserControl.Ambient.backColor
UserControl.foreColor = UserControl.Ambient.foreColor

defaultTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameDefaultLength, PropDfltDefaultLength), _
                                    PropBag.ReadProperty(PropNameDefaultUnits, PropDfltDefaultUnits))
If Err.Number <> 0 Then
    defaultTimePeriod = GetTimePeriod(PropDfltDefaultLength, PropDfltDefaultUnits)
    Err.clear
End If

Enabled = PropBag.ReadProperty(PropNameEnabled, PropDfltEnabled)
If Err.Number <> 0 Then
    Enabled = PropDfltEnabled
    Err.clear
End If

End Sub

Private Sub UserControl_Resize()
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
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameBackColor, backColor, PropDfltBackColor
PropBag.WriteProperty PropNameDefaultLength, defaultTimePeriod.length, PropDfltDefaultLength
PropBag.WriteProperty PropNameDefaultUnits, defaultTimePeriod.units, PropDfltDefaultUnits
PropBag.WriteProperty PropNameEnabled, Enabled, PropDfltEnabled
PropBag.WriteProperty PropNameForeColor, foreColor, PropDfltForeColor
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

Public Property Let backColor( _
                ByVal value As OLE_COLOR)
TimeframeLengthText.backColor = value
TimeframeUnitsCombo.backColor = value
End Property

Public Property Get backColor() As OLE_COLOR
backColor = TimeframeUnitsCombo.backColor
End Property

Public Property Get DefaultLength() As Long
DefaultLength = TimeframeLengthText
End Property

Public Property Let defaultTimePeriod( _
                ByVal value As TimePeriod)
Set mDefaultTimePeriod = value
TimeframeLengthText = value.length
setUnitsSelection value.units
End Property

Public Property Get defaultTimePeriod() As TimePeriod
Set defaultTimePeriod = mDefaultTimePeriod
End Property

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal value As Boolean)
UserControl.Enabled = value
TimeframeLengthText.Enabled = value
TimeframeUnitsCombo.Enabled = value
PropertyChanged PropNameEnabled
End Property

Public Property Let foreColor( _
                ByVal value As OLE_COLOR)
TimeframeLengthText.foreColor = value
TimeframeUnitsCombo.foreColor = value
End Property

Public Property Get foreColor() As OLE_COLOR
foreColor = TimeframeUnitsCombo.foreColor
End Property

Public Property Get isTimeframeValid() As Boolean

If TimeframeLengthText = "" Then Exit Function

If TradeBuildAPI.IsSupportedHistoricalDataPeriod(timeframeDesignator) Then isTimeframeValid = True
End Property

Public Property Get timeframeDesignator() As TimePeriod
Set timeframeDesignator = GetTimePeriod(CLng(TimeframeLengthText), TimePeriodUnitsFromString(TimeframeUnitsCombo.selectedItem.Text))
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal pTimePeriod As TimePeriod)
defaultTimePeriod = pTimePeriod
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addItem( _
                ByVal value As TimePeriodUnits)
Dim s As String
s = TimePeriodUnitsToString(value)
If TradeBuildAPI.IsSupportedHistoricalDataPeriod(GetTimePeriod(1, value)) Then TimeframeUnitsCombo.ComboItems.add , s, s
End Sub

Private Function setUnitsSelection( _
                ByVal value As TimePeriodUnits) As Boolean
If TradeBuildAPI.IsSupportedHistoricalDataPeriod(GetTimePeriod(1, value)) Then
    TimeframeUnitsCombo.ComboItems.item(TimePeriodUnitsToString(value)).Selected = True
    setUnitsSelection = True
End If
End Function

Private Sub setupTimeframeUnitsCombo()
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

End Sub
