VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.UserControl ChartNavToolbar32 
   Alignable       =   -1  'True
   ClientHeight    =   4845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10680
   ScaleHeight     =   4845
   ScaleWidth      =   10680
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   1005
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Style           =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showbars"
            Object.ToolTipText     =   "Bar chart"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showcandlesticks"
            Object.ToolTipText     =   "Candlestick chart"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showline"
            Object.ToolTipText     =   "Line chart"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showcrosshair"
            Object.ToolTipText     =   "Show crosshair"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "showdisccursor"
            Object.ToolTipText     =   "Show cursor"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "thinnerbars"
            Object.ToolTipText     =   "Thinner bars"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "thickerbars"
            Object.ToolTipText     =   "Thicker bars"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "increasespacing"
            Object.ToolTipText     =   "Increase bar spacing"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reducespacng"
            Object.ToolTipText     =   "Reduce bar spacing"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scaledown"
            Object.ToolTipText     =   "Compress vertical scale"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scaleup"
            Object.ToolTipText     =   "Expand vertical scale"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scrolldown"
            Object.ToolTipText     =   "Scroll down"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scrollup"
            Object.ToolTipText     =   "Scroll up"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scrollleft"
            Object.ToolTipText     =   "Scroll left"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scrollright"
            Object.ToolTipText     =   "Scroll right"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "scrollend"
            Object.ToolTipText     =   "Scroll to end"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "autoscale"
            Object.ToolTipText     =   "Autoscale"
            ImageIndex      =   17
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":0000
            Key             =   "showbars"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":031A
            Key             =   "showcandlesticks"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":0634
            Key             =   "showline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":094E
            Key             =   "showcrosshair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":0C68
            Key             =   "showdisccursor"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":0F82
            Key             =   "thinnerbars"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":129C
            Key             =   "thickerbars"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":15B6
            Key             =   "narrower"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":1A08
            Key             =   "wider"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":1D22
            Key             =   "scaledown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":203C
            Key             =   "scaleup"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":2356
            Key             =   "scrolldown"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":2670
            Key             =   "scrollup"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":298A
            Key             =   "scrollleft"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":2CA4
            Key             =   "scrollright"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":2FBE
            Key             =   "scrollend"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":32D8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":35F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":390C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":3C26
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":3F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":425A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":4574
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":488E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":4BA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":4FFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":544C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":5766
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":5A80
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":5D9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":60B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":63CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":66E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartNavToolbar32.ctx":6A02
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "ChartNavToolbar32"
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

Private Const ModuleName                    As String = "ChartNavToolbar32"

Private Const ChartNavCommandAutoScale           As String = "Autoscaling"

Private Const ChartNavCommandIncreaseSpacing     As String = "increasespacing"
Private Const ChartNavCommandReduceSpacing       As String = "reducespacing"

Private Const ChartNavCommandScaleDown           As String = "scaledown"
Private Const ChartNavCommandScaleUp             As String = "scaleup"

Private Const ChartNavCommandScrollDown          As String = "scrolldown"
Private Const ChartNavCommandScrollEnd           As String = "scrollend"
Private Const ChartNavCommandScrollLeft          As String = "scrollleft"
Private Const ChartNavCommandScrollRight         As String = "scrollright"
Private Const ChartNavCommandScrollUp            As String = "scrollup"

Private Const ChartNavCommandShowBars            As String = "showbars"
Private Const ChartNavCommandShowCandlesticks    As String = "showcandlesticks"
Private Const ChartNavCommandShowLine            As String = "showline"
Private Const ChartNavCommandShowCrosshair       As String = "showcrosshair"
Private Const ChartNavCommandShowDiscCursor      As String = "showdisccursor"

Private Const ChartNavCommandThickerBars         As String = "thickerbars"
Private Const ChartNavCommandThinnerBars         As String = "thinnerbars"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTradeBuildChart             As TradeBuildChart
Attribute mTradeBuildChart.VB_VarHelpID = -1
Private WithEvents mPriceRegion                 As ChartRegion
Attribute mPriceRegion.VB_VarHelpID = -1
Private mMultichartRef                          As WeakReference
Private WithEvents mBarSeries                   As BarSeries
Attribute mBarSeries.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Height = Toolbar1.Height
End Sub

Private Sub UserControl_Terminate()
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "ChartNavToolbar32 terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "ChartNavToolbar32 terminated"
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Const ProcName As String = "Toolbar1_ButtonClick"
Dim failpoint As String
On Error GoTo Err

Select Case Button.Key
Case ChartNavCommandShowBars
    mBarSeries.Style.DisplayMode = BarDisplayModeBar
Case ChartNavCommandShowCandlesticks
    mBarSeries.Style.DisplayMode = BarDisplayModeCandlestick
Case ChartNavCommandShowLine
    ' not yet implemented in ChartSkil
Case ChartNavCommandShowCrosshair
    mTradeBuildChart.PointerStyle = PointerCrosshairs
Case ChartNavCommandShowDiscCursor
    mTradeBuildChart.PointerStyle = PointerDisc
Case ChartNavCommandThinnerBars
    Select Case mBarSeries.DisplayMode
    Case BarDisplayModeCandlestick
        If mBarSeries.Width > 0.1 Then
            mBarSeries.Style.Width = mBarSeries.Width - 0.1
        End If
        If mBarSeries.Width <= 0.1 Then
            Button.Enabled = False
        End If
    Case BarDisplayModeBar
        If mBarSeries.Thickness > 1 Then
            mBarSeries.Style.Thickness = mBarSeries.Thickness - 1
        End If
        If mBarSeries.Thickness = 1 Then
            Button.Enabled = False
        End If
    End Select
Case ChartNavCommandThickerBars
    Select Case mBarSeries.DisplayMode
    Case BarDisplayModeCandlestick
        mBarSeries.Style.Width = mBarSeries.Width + 0.1
    Case BarDisplayModeBar
        mBarSeries.Style.Thickness = mBarSeries.Thickness + 1
    End Select
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = True
Case ChartNavCommandReduceSpacing
    If mTradeBuildChart.TwipsPerBar >= 50 Then
        mTradeBuildChart.TwipsPerBar = mTradeBuildChart.TwipsPerBar - 25
    End If
    If mTradeBuildChart.TwipsPerBar < 50 Then
        Button.Enabled = False
    End If
Case ChartNavCommandIncreaseSpacing
    mTradeBuildChart.TwipsPerBar = mTradeBuildChart.TwipsPerBar + 25
    Toolbar1.Buttons(ChartNavCommandReduceSpacing).Enabled = True
Case ChartNavCommandScaleDown
    mPriceRegion.ScaleUp -0.09091
Case ChartNavCommandScaleUp
    mPriceRegion.ScaleUp 0.1
Case ChartNavCommandScrollDown
    mPriceRegion.ScrollVerticalProportion -0.2
Case ChartNavCommandScrollUp
    mPriceRegion.ScrollVerticalProportion 0.2
Case ChartNavCommandScrollLeft
    mTradeBuildChart.BaseChartController.ScrollX -(mTradeBuildChart.BaseChartController.ChartWidth * 0.2)
Case ChartNavCommandScrollRight
    mTradeBuildChart.BaseChartController.ScrollX mTradeBuildChart.BaseChartController.ChartWidth * 0.2
Case ChartNavCommandScrollEnd
    mTradeBuildChart.BaseChartController.LastVisiblePeriod = mTradeBuildChart.BaseChartController.CurrentPeriodNumber
Case ChartNavCommandAutoScale
    If Button.value = tbrPressed Then
        mPriceRegion.Autoscaling = True
    Else
        mPriceRegion.Autoscaling = False
    End If
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As TWUtilities30.ChangeEventData)
Dim changeType As MultiChartChangeTypes
Const ProcName As String = "ChangeListener_Change"
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'================================================================================
' mBarSeries Event Handlers
'================================================================================

Private Sub mBarSeries_PropertyChanged(ev As TWUtilities30.PropertyChangedEventData)
Const ProcName As String = "mBarSeries_PropertyChanged"
Dim failpoint As String
On Error GoTo Err

If UCase$(ev.PropertyName) = "DISPLAYMODE" Then setupDisplayModeButtons

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mPriceRegion Event Handlers
'================================================================================

Private Sub mPriceRegion_AutoscaleChanged()
Const ProcName As String = "mPriceRegion_AutoscaleChanged"
Dim failpoint As String
On Error GoTo Err

Toolbar1.Buttons(ChartNavCommandAutoScale).value = IIf(mPriceRegion.Autoscaling, tbrPressed, tbrUnpressed)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
    setupChartNavButtons
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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
PropertyChanged "Enabled"
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                Optional ByVal pChart As TradeBuildChart, _
                Optional ByVal pMultiChart As MultiChart)
Const ProcName As String = "Initialise"
Dim failpoint As String
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
    Set mMultichartRef = CreateWeakReference(pMultiChart)
    multiChartObj.AddChangeListener Me
    attachToCurrentChart
Else
    Set mTradeBuildChart = Nothing
    Set mPriceRegion = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub attachToChart(ByVal pChart As TradeBuildChart)
Const ProcName As String = "attachToChart"
Dim failpoint As String
On Error GoTo Err

Set mTradeBuildChart = pChart
Set mPriceRegion = mTradeBuildChart.PriceRegion
Set mBarSeries = mTradeBuildChart.TradeBarSeries
If mTradeBuildChart.State = ChartStateLoaded Then setupChartNavButtons

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub attachToCurrentChart()
Const ProcName As String = "attachToCurrentChart"
Dim failpoint As String
On Error GoTo Err

Toolbar1.Enabled = False
If multiChartObj.Count > 0 Then
    Toolbar1.Enabled = True
    attachToChart multiChartObj.Chart
Else
    Toolbar1.Enabled = False
    Set mTradeBuildChart = Nothing
    Set mPriceRegion = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Function multiChartObj() As MultiChart
Const ProcName As String = "multiChartObj"
Dim failpoint As String
On Error GoTo Err

Set multiChartObj = mMultichartRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Private Sub setupChartNavButtons()

Const ProcName As String = "setupChartNavButtons"
Dim failpoint As String
On Error GoTo Err

Set mBarSeries = mTradeBuildChart.TradeBarSeries
If mBarSeries Is Nothing Then Exit Sub

setupDisplayModeButtons

If mPriceRegion.PointerStyle = PointerStyles.PointerCrosshairs Then
    Toolbar1.Buttons(ChartNavCommandShowCrosshair).value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandShowDiscCursor).value = tbrUnpressed
Else
    Toolbar1.Buttons(ChartNavCommandShowCrosshair).value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowDiscCursor).value = tbrPressed
End If

Toolbar1.Buttons(ChartNavCommandAutoScale).value = IIf(mPriceRegion.Autoscaling, tbrPressed, tbrUnpressed)

Toolbar1.Enabled = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub setupDisplayModeButtons()
Const ProcName As String = "setupDisplayModeButtons"
Dim failpoint As String
On Error GoTo Err

If mBarSeries.DisplayMode = BarDisplayModes.BarDisplayModeBar Then
    Toolbar1.Buttons(ChartNavCommandShowBars).value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandShowCandlesticks).value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = (mBarSeries.Thickness > 1)
Else
    Toolbar1.Buttons(ChartNavCommandShowBars).value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowCandlesticks).value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = (mBarSeries.Width > 0.1)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub



