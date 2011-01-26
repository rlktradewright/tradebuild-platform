VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#48.0#0"; "TWControls10.ocx"
Begin VB.UserControl ChartStylePicker 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   345
   ScaleWidth      =   1935
   Begin TWControls10.TWImageCombo ChartStylesCombo 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
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

Private Const NoChartStyle                        As String = "(None)"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTradeBuildChart                 As TradeBuildChart
Attribute mTradeBuildChart.VB_VarHelpID = -1
Private WithEvents mMultiChart                      As MultiChart
Attribute mMultiChart.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
ChartStylesCombo.Move 0, 0, UserControl.Width
UserControl.Height = ChartStylesCombo.Height
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChartStylesCombo_Click()
Const ProcName As String = "ChartStylesCombo_Click"
On Error GoTo Err

If ChartStylesCombo.Text = NoChartStyle Then
    mTradeBuildChart.BaseChartController.Style = Nothing
Else
    mTradeBuildChart.BaseChartController.Style = ChartStylesManager.item(ChartStylesCombo.Text)
End If
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mChartManager Event Handlers
'@================================================================================

Private Sub mChartManager_BaseStudyConfigurationChanged(ByVal studyConfig As ChartUtils26.StudyConfiguration)
Const ProcName As String = "mChartManager_BaseStudyConfigurationChanged"
On Error GoTo Err

selectItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mMultiChart Event Handlers
'@================================================================================

Private Sub mMultiChart_Change(ev As TWUtilities30.ChangeEventData)
Dim changeType As MultiChartChangeTypes
Const ProcName As String = "mMultiChart_Change"
Dim failpoint As String
On Error GoTo Err

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
' mTradeBuildChart Event Handlers
'@================================================================================

Private Sub mTradeBuildChart_StateChange(ev As TWUtilities30.StateChangeEventData)
Dim State As ChartStates
Const ProcName As String = "mTradeBuildChart_StateChange"
Dim failpoint As String
On Error GoTo Err

State = ev.State
Select Case State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised

Case ChartStateLoaded
    selectItem
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As String
On Error GoTo Err

UserControl.Enabled = value
ChartStylesCombo.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                Optional ByVal pChart As TradeBuildChart, _
                Optional ByVal pMultiChart As MultiChart)
Const ProcName As String = "Initialise"
On Error GoTo Err

If pChart Is Nothing And pMultiChart Is Nothing Or _
    (Not pChart Is Nothing And Not pMultiChart Is Nothing) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Either a Chart or a Multichart (but not both) must be supplied"
End If

If Not pChart Is Nothing Then
    attachToChart pChart
ElseIf Not pMultiChart Is Nothing Then
    Set mMultiChart = pMultiChart
    attachToCurrentChart
End If

loadCombo

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attachToChart(ByVal pChart As TradeBuildChart)
Const ProcName As String = "attachToChart"
Dim failpoint As String
On Error GoTo Err

Set mTradeBuildChart = pChart
If mTradeBuildChart.State = ChartStateLoaded Then selectItem

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub attachToCurrentChart()
Const ProcName As String = "attachToCurrentChart"
Dim failpoint As String
On Error GoTo Err

If mMultiChart.Count > 0 Then
    attachToChart mMultiChart.Chart
Else
    Set mTradeBuildChart = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function itemKey(ByVal pFormatterName As String, ByVal pLibraryName As String) As String
itemKey = pFormatterName & "  (" & pLibraryName & ")"
End Function

Private Sub loadCombo()
Dim lStyle As ChartStyle

Const ProcName As String = "loadCombo"
On Error GoTo Err

ChartStylesCombo.ComboItems.Clear

ChartStylesCombo.ComboItems.Add , NoChartStyle, NoChartStyle


For Each lStyle In ChartStylesManager
    ChartStylesCombo.ComboItems.Add , lStyle.name, lStyle.name
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub selectItem()
Const ProcName As String = "selectItem"
On Error GoTo Err

If mTradeBuildChart.BaseChartController.Style Is Nothing Then
    ChartStylesCombo.ComboItems.item(NoChartStyle).Selected = True
Else
    ChartStylesCombo.ComboItems.item(mTradeBuildChart.BaseChartController.Style.name).Selected = True
End If

ChartStylesCombo.SelStart = 0
ChartStylesCombo.SelLength = 0

ChartStylesCombo.Refresh

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub




