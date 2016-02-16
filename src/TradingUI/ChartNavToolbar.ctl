VERSION 5.00
Object = "{5EF6A0B6-9E1F-426C-B84A-601F4CBF70C4}#260.1#0"; "ChartSkil27.ocx"
Begin VB.UserControl ChartNavToolbar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6555
   ScaleHeight     =   3600
   ScaleWidth      =   6555
   Begin ChartSkil27.ChartToolbar ChartToolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5865
      _ExtentX        =   9737
      _ExtentY        =   582
   End
End
Attribute VB_Name = "ChartNavToolbar"
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

Implements IChangeListener
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

Private Const ModuleName                    As String = "ChartNavToolbar"


'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTradeBuildChart             As MarketChart
Attribute mTradeBuildChart.VB_VarHelpID = -1
Private WithEvents mChartManager                As ChartManager
Attribute mChartManager.VB_VarHelpID = -1
Private mMultichartRef                          As WeakReference

Private mTheme                                  As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Height = ChartToolbar1.Height
UserControl.Width = ChartToolbar1.Width
End Sub

Private Sub UserControl_Terminate()
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "ChartNavToolbar terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "ChartNavToolbar terminated"
End Sub

'@================================================================================
' IChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "IChangeListener_Change"
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
gHandleUnexpectedError ProcName, ModuleName
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

'================================================================================
' Control Event Handlers
'================================================================================

'@================================================================================
' mChartManager Event Handlers
'@================================================================================

Private Sub mChartManager_BaseStudyConfigurationChanged(ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "mChartManager_BaseStudyConfigurationChanged"
On Error GoTo Err

setupChartNavButtons

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTradeBuildChart Event Handlers
'@================================================================================

Private Sub mTradeBuildChart_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTradeBuildChart_StateChange"
On Error GoTo Err

Dim State As ChartStates

State = ev.State
Select Case State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised

Case ChartStateLoaded
    setChartManager
    setupChartNavButtons
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
ChartToolbar1.Enabled = Value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.ToolbarBackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
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

If Not pChart Is Nothing Then
    attachToChart pChart
ElseIf Not pMultiChart Is Nothing Then
    Set mMultichartRef = CreateWeakReference(pMultiChart)
    multiChartObj.AddChangeListener Me
    attachToCurrentChart
Else
    Set mTradeBuildChart = Nothing
End If

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

    Set mTradeBuildChart = pChart
    If mTradeBuildChart.State = ChartStateLoaded Then
        setChartManager
        setupChartNavButtons
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub attachToCurrentChart()
Const ProcName As String = "attachToCurrentChart"
On Error GoTo Err

If multiChartObj.Count > 0 Then
    attachToChart multiChartObj.Chart
Else
    Set mTradeBuildChart = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function multiChartObj() As MultiChart
Const ProcName As String = "multiChartObj"
On Error GoTo Err

Set multiChartObj = mMultichartRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setChartManager()
Set mChartManager = mTradeBuildChart.ChartManager
End Sub

Private Sub setupChartNavButtons()
Const ProcName As String = "setupChartNavButtons"
On Error GoTo Err

ChartToolbar1.Initialise mTradeBuildChart.BaseChartController, _
                        mTradeBuildChart.PriceRegion, _
                        mTradeBuildChart.TradeBarSeries

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

