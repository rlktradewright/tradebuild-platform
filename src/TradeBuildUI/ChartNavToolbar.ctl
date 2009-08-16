VERSION 5.00
Object = "{74951842-2BEF-4829-A34F-DC7795A37167}#115.1#0"; "ChartSkil2-6.ocx"
Begin VB.UserControl ChartNavToolbar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   ScaleHeight     =   3600
   ScaleWidth      =   6945
   Begin ChartSkil26.ChartToolbar ChartToolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
      _ExtentX        =   12250
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

Implements ChangeListener

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

Private WithEvents mTradeBuildChart             As TradeBuildChart
Attribute mTradeBuildChart.VB_VarHelpID = -1
Private mMultichartRef                          As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Height = ChartToolbar1.Height
End Sub

Private Sub UserControl_Terminate()
gLogger.Log LogLevelDetail, "ChartNavToolbar terminated"
Debug.Print "ChartNavToolbar terminated"
End Sub

'================================================================================
' Control Event Handlers
'================================================================================


'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As TWUtilities30.ChangeEvent)
Dim changeType As MultiChartChangeTypes
changeType = ev.changeType
Select Case changeType
Case MultiChartSelectionChanged
    attachToCurrentChart
Case MultiChartAdd

Case MultiChartRemove

End Select
End Sub

'@================================================================================
' mTradeBuildChart Event Handlers
'@================================================================================

Private Sub mTradeBuildChart_StateChange(ev As TWUtilities30.StateChangeEvent)
Dim State As ChartStates
State = ev.State
Select Case State
Case ChartStateBlank

Case ChartStateCreated

Case ChartStateInitialised

Case ChartStateLoaded
    setupChartNavButtons
End Select
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
UserControl.Enabled = value
ChartToolbar1.Enabled = value
PropertyChanged "Enabled"
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                Optional ByVal pChart As TradeBuildChart, _
                Optional ByVal pMultiChart As MultiChart)
If pChart Is Nothing And pMultiChart Is Nothing Or _
    (Not pChart Is Nothing And Not pMultiChart Is Nothing) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "initialise", _
            "Either a Chart or a Multichart (but not both) must be supplied"
End If

If Not pChart Is Nothing Then
    attachToChart pChart
ElseIf Not pMultiChart Is Nothing Then
    Set mMultichartRef = CreateWeakReference(pMultiChart)
    multiChartObj.AddChangeListener Me
    attachToCurrentChart
Else
    Set mTradeBuildChart = Nothing
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attachToChart(ByVal pChart As TradeBuildChart)
    Set mTradeBuildChart = pChart
    If mTradeBuildChart.State = ChartStateLoaded Then setupChartNavButtons
End Sub

Private Sub attachToCurrentChart()
If multiChartObj.count > 0 Then
    attachToChart multiChartObj.Chart
Else
    Set mTradeBuildChart = Nothing
End If
End Sub

Private Function multiChartObj() As MultiChart
Set multiChartObj = mMultichartRef.Target
End Function

Private Sub setupChartNavButtons()

ChartToolbar1.Initialise mTradeBuildChart.BaseChartController, _
                        mTradeBuildChart.PriceRegion, _
                        mTradeBuildChart.TradeBarSeries

End Sub

