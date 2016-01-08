VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl ChartStylePicker 
   BackStyle       =   0  'Transparent
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   345
   ScaleWidth      =   1935
   Begin TWControls40.TWImageCombo Combo1 
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
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
      MouseIcon       =   "ChartStylePicker.ctx":0000
      Text            =   ""
   End
End
Attribute VB_Name = "ChartStylePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ChartStylePicker"

Private Const NoChartStyle                          As String = "(None)"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mMarketChart                     As MarketChart
Attribute mMarketChart.VB_VarHelpID = -1
Private WithEvents mMultiChart                      As MultiChart
Attribute mMultiChart.VB_VarHelpID = -1

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
Combo1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
Combo1.ListWidth = PropBag.ReadProperty("ListWidth", 0)
Combo1.Locked = PropBag.ReadProperty("Locked", False)
Combo1.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

Combo1.Move 0, 0, UserControl.Width
UserControl.Height = Combo1.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
Call PropBag.WriteProperty("CausesValidation", Combo1.CausesValidation, True)
Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
Call PropBag.WriteProperty("ListWidth", Combo1.ListWidth, 0)
Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
Call PropBag.WriteProperty("ToolTipText", Combo1.ToolTipText, "")
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

Private Sub Combo1_Click()
Const ProcName As String = "Combo1_Click"
On Error GoTo Err

If Combo1.Text = NoChartStyle Then
    mMarketChart.BaseChartController.Style = Nothing
Else
    mMarketChart.BaseChartController.Style = ChartStylesManager.Item(Combo1.Text)
End If
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mChartManager Event Handlers
'@================================================================================

Private Sub mChartManager_BaseStudyConfigurationChanged(ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "mChartManager_BaseStudyConfigurationChanged"
On Error GoTo Err

SelectItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mMultiChart Event Handlers
'@================================================================================

Private Sub mMultiChart_Change(ev As ChangeEventData)
Const ProcName As String = "mMultiChart_Change"
On Error GoTo Err

Dim changeType As MultiChartChangeTypes

changeType = ev.changeType
Select Case changeType
Case MultiChartSelectionChanged
    attachToCurrentChart
Case MultiChartAdd

Case MultiChartRemove

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mMarketChart Event Handlers
'@================================================================================

Private Sub mMarketChart_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mMarketChart_StateChange"
On Error GoTo Err

Dim State As ChartStates

State = ev.State
Select Case State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised

Case ChartStateLoaded
    SelectItem
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mMarketChart_StyleChanged(ByVal pNewStyle As ChartStyle)
Const ProcName As String = "mMarketChart_StyleChanged"
On Error GoTo Err

SelectStyle pNewStyle.Name

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
Const ProcName As String = "BackColor"
On Error GoTo Err

BackColor = Combo1.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
Const ProcName As String = "BackColor"
On Error GoTo Err

Combo1.BackColor() = New_BackColor
PropertyChanged "BackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CausesValidation() As Boolean
Const ProcName As String = "CausesValidation"
On Error GoTo Err

CausesValidation = Combo1.CausesValidation

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
Const ProcName As String = "CausesValidation"
On Error GoTo Err

Combo1.CausesValidation() = New_CausesValidation
PropertyChanged "CausesValidation"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = value
Combo1.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
Const ProcName As String = "Font"
On Error GoTo Err

Set Font = Combo1.Font

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set Font(ByVal New_Font As Font)
Const ProcName As String = "Font"
On Error GoTo Err

Set Combo1.Font = New_Font
PropertyChanged "Font"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
Const ProcName As String = "ForeColor"
On Error GoTo Err

ForeColor = Combo1.ForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

Combo1.ForeColor() = New_ForeColor
PropertyChanged "ForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Locked() As Boolean
Const ProcName As String = "Locked"
On Error GoTo Err

Locked = Combo1.Locked

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
Const ProcName As String = "Locked"
On Error GoTo Err

Combo1.Locked() = New_Locked
PropertyChanged "Locked"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ListWidth(ByVal value As Long)
Const ProcName As String = "ListWidth"
On Error GoTo Err

Combo1.ListWidth = value
PropertyChanged "ListWidth"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ListWidth() As Long
ListWidth = Combo1.ListWidth
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = value
If mTheme Is Nothing Then Exit Property

Combo1.Theme = mTheme

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Get ToolTipText() As String
Const ProcName As String = "ToolTipText"
On Error GoTo Err

ToolTipText = Combo1.ToolTipText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
Const ProcName As String = "ToolTipText"
On Error GoTo Err

Combo1.ToolTipText() = New_ToolTipText
PropertyChanged "ToolTipText"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                Optional ByVal pChart As MarketChart, _
                Optional ByVal pMultiChart As MultiChart)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument (Not pChart Is Nothing Or Not pMultiChart Is Nothing) And _
    (pChart Is Nothing Or pMultiChart Is Nothing), _
    "Either a Chart or a Multichart (but not both) must be supplied"


loadCombo

If Not pChart Is Nothing Then
    attachToChart pChart
ElseIf Not pMultiChart Is Nothing Then
    Set mMultiChart = pMultiChart
    attachToCurrentChart
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectStyle(ByVal pStyleName As String)
Const ProcName As String = "SelectStyle"
On Error GoTo Err

Set Combo1.SelectedItem = Combo1.ComboItems.Item(pStyleName)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attachToChart(ByVal pChart As MarketChart)
Const ProcName As String = "attachToChart"
On Error GoTo Err

Set mMarketChart = pChart
If mMarketChart.State = ChartStateLoaded Then SelectItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub attachToCurrentChart()
Const ProcName As String = "attachToCurrentChart"
On Error GoTo Err

If mMultiChart.Count > 0 Then
    attachToChart mMultiChart.Chart
Else
    Set mMarketChart = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadCombo()
Const ProcName As String = "loadCombo"
On Error GoTo Err

Dim lStyle As ChartStyle

Combo1.ComboItems.Clear

Combo1.ComboItems.Add , NoChartStyle, NoChartStyle


For Each lStyle In ChartStylesManager
    Combo1.ComboItems.Add , lStyle.Name, lStyle.Name
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub SelectItem()
Const ProcName As String = "selectItem"
On Error GoTo Err

If mMarketChart.BaseChartController.Style Is Nothing Then
    SelectStyle NoChartStyle
Else
    SelectStyle mMarketChart.BaseChartController.Style.Name
End If

Combo1.SelStart = 0
Combo1.SelLength = 0

Combo1.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




