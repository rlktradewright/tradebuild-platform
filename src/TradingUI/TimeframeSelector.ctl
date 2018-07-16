VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.UserControl TimeframeSelector 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3705
   ScaleHeight     =   1710
   ScaleWidth      =   3705
   Begin TWControls40.TWImageCombo TimeframeCombo 
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
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
      MouseIcon       =   "TimeframeSelector.ctx":0000
      Text            =   ""
   End
End
Attribute VB_Name = "TimeframeSelector"
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
Event Click()
Attribute Click.VB_UserMemId = -600

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "TimeframeSelector"

Private Const PropNameBackcolor                         As String = "BackColor"
Private Const PropNameForecolor                         As String = "ForeColor"

Private Const PropDfltBackColor                         As Long = vbWindowBackground
Private Const PropDfltForeColor                         As Long = vbWindowText

Private Const TimeframeCustom                           As String = "Custom"

'@================================================================================
' Member variables
'@================================================================================

Private mSpecifier                                      As fTimeframeSpecifier

Private mLatestTimePeriod                               As TimePeriod

Private mValidator                                      As ITimePeriodValidator

Private mTheme                                          As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_InitProperties()
On Error Resume Next

BackColor = PropDfltBackColor
ForeColor = PropDfltForeColor
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

BackColor = PropBag.ReadProperty(PropNameBackcolor, PropDfltBackColor)
ForeColor = PropBag.ReadProperty(PropNameForecolor, PropDfltForeColor)

End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

TimeframeCombo.Left = 0
TimeframeCombo.Top = 0
TimeframeCombo.Width = UserControl.Width
UserControl.Height = TimeframeCombo.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Terminate()
Const ProcName As String = "UserControl_Terminate"
On Error GoTo Err

If Not mSpecifier Is Nothing Then Unload mSpecifier

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next
PropBag.WriteProperty PropNameBackcolor, BackColor, PropDfltBackColor
PropBag.WriteProperty PropNameForecolor, ForeColor, PropDfltForeColor
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

Private Sub TimeframeCombo_Change()
RaiseEvent Change
End Sub

Private Sub TimeframeCombo_Click()
Const ProcName As String = "TimeframeCombo_Click"
On Error GoTo Err

If TimeframeCombo.Text = TimeframeCustom Then
    If mSpecifier Is Nothing Then
        Set mSpecifier = New fTimeframeSpecifier
        mSpecifier.Theme = mTheme
        mSpecifier.Initialise mValidator
    End If
    
    Dim myPosition As GDI_POINT
    myPosition.X = 0
    myPosition.Y = UserControl.Height / Screen.TwipsPerPixelY
    MapWindowPoints UserControl.hWnd, 0, VarPtr(myPosition), 1
    
    Dim X As Long: X = myPosition.X * Screen.TwipsPerPixelX
    If X + mSpecifier.Width > Screen.Width Then X = Screen.Width - mSpecifier.Width
    
    Dim Y As Long: Y = myPosition.Y * Screen.TwipsPerPixelY
    If Y + mSpecifier.Height > Screen.Height Then Y = Y - UserControl.Height - mSpecifier.Height
    
    mSpecifier.Move X, Y
    mSpecifier.ZOrder 0
    mSpecifier.Show vbModal, gGetParentForm(Me)
    
    
    If Not mSpecifier.Cancelled Then
        Dim tp As TimePeriod
        Set tp = mSpecifier.TimePeriod
        selectComboEntry tp
        RaiseEvent Click
    Else
        selectComboEntry mLatestTimePeriod
    End If
Else
    RaiseEvent Click
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

TimeframeCombo.BackColor = Value
PropertyChanged PropNameBackcolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
Const ProcName As String = "backColor"
On Error GoTo Err

BackColor = TimeframeCombo.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled( _
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "foreColor"
On Error GoTo Err

TimeframeCombo.ForeColor = Value
PropertyChanged PropNameForecolor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
Const ProcName As String = "foreColor"
On Error GoTo Err

ForeColor = TimeframeCombo.ForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Const ProcName As String = "Text"
On Error GoTo Err

Text = TimeframeCombo.Text

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

BackColor = mTheme.TextBackColor
ForeColor = mTheme.TextForeColor
If Not mSpecifier Is Nothing Then mSpecifier.Theme = mTheme

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let TimePeriod( _
                ByRef Value As TimePeriod)
Const ProcName As String = "TimePeriod"
On Error GoTo Err

selectComboEntry Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod() As TimePeriod
Const ProcName As String = "TimePeriod"
On Error GoTo Err

If TimeframeCombo.SelectedItem Is Nothing Then Exit Property

Set TimePeriod = TimePeriodFromString(TimeframeCombo.SelectedItem.Text)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AppendEntry( _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "AppendEntry"
On Error GoTo Err

addComboEntry pTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise(ByVal pValidator As ITimePeriodValidator)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pValidator Is Nothing, "pValidator cannot be Nothing"

Set mValidator = pValidator

TimeframeCombo.ComboItems.Clear
TimeframeCombo.ComboItems.Add , TimeframeCustom, TimeframeCustom
addComboEntry GetTimePeriod(5, TimePeriodSecond)
addComboEntry GetTimePeriod(15, TimePeriodSecond)
addComboEntry GetTimePeriod(30, TimePeriodSecond)
addComboEntry GetTimePeriod(1, TimePeriodMinute)
addComboEntry GetTimePeriod(2, TimePeriodMinute)
addComboEntry GetTimePeriod(3, TimePeriodMinute)
addComboEntry GetTimePeriod(4, TimePeriodMinute)
addComboEntry GetTimePeriod(5, TimePeriodMinute)
addComboEntry GetTimePeriod(8, TimePeriodMinute)
addComboEntry GetTimePeriod(10, TimePeriodMinute)
addComboEntry GetTimePeriod(13, TimePeriodMinute)
addComboEntry GetTimePeriod(15, TimePeriodMinute)
addComboEntry GetTimePeriod(20, TimePeriodMinute)
addComboEntry GetTimePeriod(21, TimePeriodMinute)
addComboEntry GetTimePeriod(30, TimePeriodMinute)
addComboEntry GetTimePeriod(34, TimePeriodMinute)
addComboEntry GetTimePeriod(55, TimePeriodMinute)
addComboEntry GetTimePeriod(1, TimePeriodHour)
addComboEntry GetTimePeriod(1, TimePeriodDay)
addComboEntry GetTimePeriod(1, TimePeriodWeek)
addComboEntry GetTimePeriod(1, TimePeriodMonth)
addComboEntry GetTimePeriod(100, TimePeriodVolume)
addComboEntry GetTimePeriod(1000, TimePeriodVolume)
addComboEntry GetTimePeriod(4, TimePeriodTickMovement)
addComboEntry GetTimePeriod(5, TimePeriodTickMovement)
addComboEntry GetTimePeriod(10, TimePeriodTickMovement)
addComboEntry GetTimePeriod(20, TimePeriodTickMovement)

'If selectComboEntry(5, TimePeriodMinute) Then
'ElseIf selectComboEntry(10, TimePeriodMinute) Then
'ElseIf selectComboEntry(15, TimePeriodMinute) Then
'ElseIf selectComboEntry(20, TimePeriodMinute) Then
'ElseIf selectComboEntry(30, TimePeriodMinute) Then
'ElseIf selectComboEntry(1, TimePeriodHour) Then
'ElseIf selectComboEntry(1, TimePeriodDay) Then
'ElseIf selectComboEntry(1, TimePeriodWeek) Then
'Else
'    selectComboEntry 1, TimePeriodMonth
'End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub InsertEntry( _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "insertEntry"
On Error GoTo Err

insertComboEntry pTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectTimeframe( _
                ByRef pTimePeriod As TimePeriod)
Const ProcName As String = "selectTimeframe"
On Error GoTo Err

selectComboEntry pTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addComboEntry( _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "addComboEntry"
On Error GoTo Err

Dim s As String

s = pTimePeriod.ToString
If Not mValidator.IsValidTimePeriod(pTimePeriod) Then Exit Sub
TimeframeCombo.ComboItems.Add , s, s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub insertComboEntry( _
                ByVal pTimePeriod As TimePeriod)
Const ProcName As String = "insertComboEntry"
On Error GoTo Err

Dim s As String
Dim i As Long
Dim unitFound As Boolean

s = pTimePeriod.ToString
If mValidator.IsValidTimePeriod(pTimePeriod) Then
    For i = 2 To TimeframeCombo.ComboItems.Count
        Dim currTp As TimePeriod
        Set currTp = TimePeriodFromString(TimeframeCombo.ComboItems(i).Text)
        If currTp.Units = pTimePeriod.Units Then unitFound = True
        If currTp.Units = pTimePeriod.Units And currTp.Length >= pTimePeriod.Length Then Exit For
        If unitFound And currTp.Units <> pTimePeriod.Units Then Exit For
    Next
    If currTp.Units = pTimePeriod.Units And currTp.Length = pTimePeriod.Length Then Exit Sub
    TimeframeCombo.ComboItems.Add i, s, s
    TimeframeCombo.Refresh
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function selectComboEntry( _
                ByVal pTimePeriod As TimePeriod) As Boolean
Const ProcName As String = "selectComboEntry"
On Error GoTo Err

Dim s As String
s = pTimePeriod.ToString

If mValidator.IsValidTimePeriod(pTimePeriod) Then
    Set TimeframeCombo.SelectedItem = TimeframeCombo.ComboItems.Item(s)
    Set mLatestTimePeriod = pTimePeriod
    selectComboEntry = True
End If

Exit Function

Err:

If Err.Number = 35601 Then
    insertComboEntry pTimePeriod
    Resume
End If
gHandleUnexpectedError ProcName, ModuleName
End Function


