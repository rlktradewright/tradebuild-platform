VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.UserControl ChartToolbar 
   Alignable       =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7425
   ScaleHeight     =   3600
   ScaleWidth      =   7425
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   600
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":2B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":2F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":33D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":3C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":40CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":4520
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":4972
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":4DC4
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":5216
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":5668
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":5ABA
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":5F0C
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":635E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":67B0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":6C02
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":7054
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":74A6
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":78F8
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":7D4A
            Key             =   "IMG13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":819C
            Key             =   "IMG14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":85EE
            Key             =   "IMG15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":8A40
            Key             =   "IMG16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartToolbar.ctx":8E92
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList4"
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
            Key             =   "reducespacing"
            Object.ToolTipText     =   "Reduce bar spacing"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "increasespacing"
            Object.ToolTipText     =   "Increase bar spacing"
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
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "ChartToolbar"
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

Private Const ModuleName                        As String = "ChartToolbar"

Private Const ChartNavCommandAutoScale           As String = "autoscale"
Private Const ChartNavCommandAutoScroll          As String = "autoscroll"

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
Private Const ChartNavCommandShowPlainCursor     As String = "showplaincursor"
Private Const ChartNavCommandShowDiscCursor      As String = "showdisccursor"

Private Const ChartNavCommandThickerBars         As String = "thickerbars"
Private Const ChartNavCommandThinnerBars         As String = "thinnerbars"

'@================================================================================
' Member variables
'@================================================================================

Private mController                             As ChartController
Attribute mController.VB_VarHelpID = -1
Private WithEvents mRegion                      As ChartRegion
Attribute mRegion.VB_VarHelpID = -1
Private mMultichartRef                          As WeakReference
Private WithEvents mBarSeries                   As BarSeries
Attribute mBarSeries.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
UserControl.Height = Toolbar1.Height
UserControl.Width = Toolbar1.Width
End Sub

Private Sub UserControl_Terminate()
'gLogger.Log LogLevelDetail, "ChartToolbar terminated"
Debug.Print "ChartToolbar terminated"
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Const ProcName As String = "Toolbar1_ButtonClick"

On Error GoTo Err

Select Case Button.Key
Case ChartNavCommandShowBars
    mBarSeries.Style.DisplayMode = BarDisplayModeBar
Case ChartNavCommandShowCandlesticks
    mBarSeries.Style.DisplayMode = BarDisplayModeCandlestick
Case ChartNavCommandShowLine
    mBarSeries.Style.DisplayMode = BarDisplayModeLine
Case ChartNavCommandShowCrosshair
    mController.PointerStyle = PointerCrosshairs
Case ChartNavCommandShowDiscCursor
    mController.PointerStyle = PointerDisc
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
    Case BarDisplayModeLine
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
    Case BarDisplayModeLine
        mBarSeries.Style.Thickness = mBarSeries.Thickness + 1
    End Select
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = True
Case ChartNavCommandReduceSpacing
    If mController.TwipsPerPeriod >= 50 Then
        mController.TwipsPerPeriod = mController.TwipsPerPeriod - 25
    End If
    If mController.TwipsPerPeriod < 50 Then
        Button.Enabled = False
    End If
Case ChartNavCommandIncreaseSpacing
    mController.TwipsPerPeriod = mController.TwipsPerPeriod + 25
    Toolbar1.Buttons(ChartNavCommandReduceSpacing).Enabled = True
Case ChartNavCommandScaleDown
    mRegion.ScaleUp -0.09091
Case ChartNavCommandScaleUp
    mRegion.ScaleUp 0.1
Case ChartNavCommandScrollDown
    mRegion.ScrollVerticalProportion -0.2
Case ChartNavCommandScrollUp
    mRegion.ScrollVerticalProportion 0.2
Case ChartNavCommandScrollLeft
    mController.ScrollX -(mController.ChartWidth * 0.2)
Case ChartNavCommandScrollRight
    mController.ScrollX mController.ChartWidth * 0.2
Case ChartNavCommandScrollEnd
    mController.LastVisiblePeriod = mController.CurrentPeriodNumber
Case ChartNavCommandAutoScale
    If Button.Value = tbrPressed Then
        mRegion.Autoscaling = True
    Else
        mRegion.Autoscaling = False
    End If
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mBarSeries Event Handlers
'================================================================================

Private Sub mBarSeries_PropertyChanged(ev As PropertyChangedEventData)
Const ProcName As String = "mBarSeries_PropertyChanged"

On Error GoTo Err

Select Case UCase$(ev.PropertyName)
Case "DISPLAYMODE"
    setupDisplayModeButtons
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'================================================================================
' mRegion Event Handlers
'================================================================================

Private Sub mRegion_AutoscalingChanged()
Const ProcName As String = "mRegion_AutoscalingChanged"

On Error GoTo Err

Toolbar1.Buttons(ChartNavCommandAutoScale).Value = IIf(mRegion.Autoscaling, tbrPressed, tbrUnpressed)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"

On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Enabled( _
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"

On Error GoTo Err

UserControl.Enabled = Value
If Not mController Is Nothing And _
    Not mRegion Is Nothing And _
    Not mBarSeries Is Nothing _
Then
    Toolbar1.Enabled = Value
Else
    Toolbar1.Enabled = False
End If
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pChartController As ChartController, _
                ByVal pRegion As ChartRegion, _
                ByVal pBarSeries As BarSeries)

Const ProcName As String = "Initialise"

On Error GoTo Err

If Not pChartController Is Nothing And _
    Not pRegion Is Nothing And _
    Not pBarSeries Is Nothing _
Then
    Set mController = pChartController
    Set mRegion = pRegion
    Set mBarSeries = pBarSeries
    setupChartNavButtons
    Toolbar1.Enabled = True
Else
    Set mController = Nothing
    Set mRegion = Nothing
    Set mBarSeries = Nothing
    Toolbar1.Enabled = False
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setupChartNavButtons()

Const ProcName As String = "setupChartNavButtons"

On Error GoTo Err

If mBarSeries Is Nothing Then Exit Sub

setupDisplayModeButtons

If mRegion.PointerStyle = PointerStyles.PointerCrosshairs Then
    Toolbar1.Buttons(ChartNavCommandShowCrosshair).Value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandShowDiscCursor).Value = tbrUnpressed
Else
    Toolbar1.Buttons(ChartNavCommandShowCrosshair).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowDiscCursor).Value = tbrPressed
End If

Toolbar1.Buttons(ChartNavCommandAutoScale).Value = IIf(mRegion.Autoscaling, tbrPressed, tbrUnpressed)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub setupDisplayModeButtons()
Const ProcName As String = "setupDisplayModeButtons"

On Error GoTo Err

If mBarSeries.DisplayMode = BarDisplayModes.BarDisplayModeBar Then
    Toolbar1.Buttons(ChartNavCommandShowBars).Value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandShowCandlesticks).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowLine).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = (mBarSeries.Thickness > 1)
ElseIf mBarSeries.DisplayMode = BarDisplayModes.BarDisplayModeCandlestick Then
    Toolbar1.Buttons(ChartNavCommandShowBars).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowCandlesticks).Value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandShowLine).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = (mBarSeries.Width > 0.1)
Else
    Toolbar1.Buttons(ChartNavCommandShowBars).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowCandlesticks).Value = tbrUnpressed
    Toolbar1.Buttons(ChartNavCommandShowLine).Value = tbrPressed
    Toolbar1.Buttons(ChartNavCommandThinnerBars).Enabled = (mBarSeries.Thickness > 1)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

