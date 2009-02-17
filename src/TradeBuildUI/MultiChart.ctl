VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl MultiChart 
   Alignable       =   -1  'True
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   7140
   ScaleWidth      =   9480
   Begin TradeBuildUI26.TradeBuildChart TBChart 
      Align           =   1  'Align Top
      Height          =   5415
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   9551
      ChartBackColor  =   6566450
      ShowToobar      =   0   'False
   End
   Begin MSComctlLib.Toolbar ControlToolbar 
      Height          =   330
      Left            =   6480
      TabIndex        =   0
      Top             =   6480
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "selecttimeframe"
            Object.ToolTipText     =   "Choose the timeframe for the new chart"
            Style           =   4
            Object.Width           =   1700
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Select a new timeframe and add another chart"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Object.ToolTipText     =   "Remove current chart"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin TradeBuildUI26.TimeframeSelector TimeframeSelector1 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Choose the timeframe for the new chart"
         Top             =   0
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2400
      Top             =   5880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":0452
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip ChartSelector 
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   6555
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   582
      Style           =   1
      HotTracking     =   -1  'True
      Placement       =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MultiChart"
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

Event Change(ev As ChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "MultiChart"

'@================================================================================
' Member variables
'@================================================================================

Private mTicker                             As ticker
Private mSpec                               As ChartSpecifier
Private mIsHistoric                         As Boolean
Private mFromTime                           As Date
Private mToTime                             As Date

Private mIndexes                            As Collection
Private mCurrentIndex                       As Long

Private mBarFormatterFactory                As BarFormatterFactory

Private mChangeListeners                    As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mIndexes = New Collection
Set mChangeListeners = New Collection
ChartSelector.Tabs.clear
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
hideTimeframeSelector
End Sub

Private Sub UserControl_Resize()
resize
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChartSelector_Click()
switchToChart ChartSelector.selectedItem.index
End Sub

Private Sub ControlToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
Case "add"
    If ControlToolbar.Buttons("add").value = tbrPressed Then
        showTimeframeSelector
    Else
        hideTimeframeSelector
    End If
Case "remove"
    remove mCurrentIndex
End Select
End Sub

Private Sub TBChart_TimeframeChange(index As Integer)
ChartSelector.Tabs(getIndexFromChartControlIndex(index)).caption = TBChart(index).TimePeriod.toShortString
fireChange MultiChartTimeframeChanged
End Sub

Private Sub TimeframeSelector1_Click()
add TimeframeSelector1.timeframeDesignator
hideTimeframeSelector
ControlToolbar.Buttons("add").value = tbrUnpressed
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get chartController( _
                Optional ByVal index As Long = -1) As chartController
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set chartController = getChartFromIndex(index).chartController

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "chartController" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get chartManager( _
                Optional ByVal index As Long = -1) As chartManager
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set chartManager = getChartFromIndex(index).chartManager

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "chartManager" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get count() As Long
count = ChartSelector.Tabs.count
End Property

Public Property Get Chart( _
                Optional ByVal index As Long = -1) As TradeBuildChart
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set Chart = getChartFromIndex(index)

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Chart" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
UserControl.Enabled = value
PropertyChanged "Enabled"
End Property

Public Property Get priceRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set priceRegion = getChartFromIndex(index).priceRegion

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "priceRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get State( _
                Optional ByVal index As Long = -1) As ChartStates
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
State = getChartFromIndex(index).State

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "state" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Timeframe( _
                Optional ByVal index As Long = -1) As Timeframe
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set Timeframe = getChartFromIndex(index).Timeframe

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Timeframe" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get TimePeriod( _
                Optional ByVal index As Long = -1) As TimePeriod
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set TimePeriod = getChartFromIndex(index).TimePeriod

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "TimePeriod" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get volumeRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set volumeRegion = getChartFromIndex(index).volumeRegion

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "volumeRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function add( _
                ByVal Timeframe As TimePeriod) As TradeBuildChart
Dim lChart As TradeBuildChart
Dim failpoint As Long
On Error GoTo Err

load TBChart(TBChart.UBound + 1)
Set lChart = TBChart(TBChart.UBound).Object
TBChart(TBChart.UBound).align = vbAlignTop
TBChart(TBChart.UBound).Top = 0
TBChart(TBChart.UBound).Height = ChartSelector.Top
mSpec.Timeframe = Timeframe
If mIsHistoric Then
    lChart.showHistoricChart mTicker, mSpec, mFromTime, mToTime, mBarFormatterFactory
Else
    lChart.showChart mTicker, mSpec, mBarFormatterFactory
End If

addTab lChart

fireChange MultiChartAdd
fireChange MultiChartSelectionChanged

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "add" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub AddChangeListener( _
                ByVal listener As ChangeListener)
mChangeListeners.add listener, CStr(ObjPtr(listener))
End Sub
               
Public Sub ChangeTimeframe(ByVal pTimeframe As TimePeriod)
Dim failpoint As Long
On Error GoTo Err

Chart.ChangeTimeframe pTimeframe

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChangeTimeframe" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub clear()
Do While ChartSelector.Tabs.count <> 0
    remove ChartSelector.Tabs.count
Loop
End Sub

Public Sub finish()
Dim i As Long
Dim index As Long
For i = 1 To mIndexes.count
    index = mIndexes(i)
    getChartFromIndex(index).finish
    Unload TBChart(getChartControlIndexFromIndex(index))
Next
End Sub

Public Sub Initialise( _
                ByVal pTicker As ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal fromTime As Date, _
                Optional ByVal toTime As Date, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)
Dim failpoint As Long
On Error GoTo Err

Set mTicker = pTicker
Set mSpec = chartSpec
TBChart(0).ChartBackGradientFillColors = mSpec.ChartBackGradientFillColors
mIsHistoric = (fromTime <> 0 Or toTime <> 0)
mFromTime = fromTime
mToTime = toTime
Set mBarFormatterFactory = BarFormatterFactory
TimeframeSelector1.Initialise
TimeframeSelector1.selectTimeframe mSpec.Timeframe

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Initialise" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub remove( _
                ByVal index As Long)
Dim nxtIndex As Long
Dim failpoint As Long
On Error GoTo Err

If index > count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Remove", _
            "Index must not be less than 1 or greater than Count"
End If

nxtIndex = nextIndex(index)
closeChart index
If index = mCurrentIndex Then mCurrentIndex = 0
If nxtIndex <> 0 Then SelectChart nxtIndex

fireChange MultiChartRemove
fireChange MultiChartSelectionChanged

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "remove" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveChangeListener(ByVal listener As ChangeListener)
mChangeListeners.remove ObjPtr(listener)
End Sub

Public Sub scrollToTime(ByVal pTime As Date)
Dim failpoint As Long
On Error GoTo Err

Chart.scrollToTime pTime

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "scrollToTime" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SelectChart( _
                ByVal index As Long)
Dim failpoint As Long
On Error GoTo Err

If index > count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SelectChart", _
            "Index must not be less than 1 or greater than Count"
End If

ChartSelector.Tabs(index).Selected = True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SelectChart" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addTab( _
                ByVal pChart As TradeBuildChart)
Static chartNumber As Long
Dim lTab As MSComctlLib.Tab

chartNumber = chartNumber + 1
Set lTab = ChartSelector.Tabs.add(, , pChart.TimePeriod.toShortString)
lTab.Tag = CStr(chartNumber)

ControlToolbar.Buttons("remove").Enabled = True

mIndexes.add TBChart.UBound, CStr(chartNumber)
lTab.Selected = True
End Sub

Private Function checkIndex( _
                ByVal index As Long) As Long
If index = -1 Then
    If mCurrentIndex < 1 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & "checkIndex", _
                "No currrent chart"
    Else
        index = mCurrentIndex
    End If
End If

If index > count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "checkIndex", _
            "Index must not be less than 1 or greater than Count"
End If

checkIndex = index
End Function

Private Sub closeChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
Set lChart = getChartFromIndex(index)
lChart.finish
Unload TBChart(getChartControlIndexFromIndex(index))
mIndexes.remove ChartSelector.Tabs(index).Tag
ChartSelector.Tabs.remove index
If ChartSelector.Tabs.count = 0 Then
    ControlToolbar.Buttons("remove").Enabled = False
    TBChart(0).Visible = True
    TBChart(0).Top = 0
    TBChart(0).Height = ChartSelector.Top
End If
End Sub

Private Sub fireChange( _
                ByVal changeType As MultiChartChangeTypes)
Dim listener As ChangeListener
Dim ev As ChangeEvent
Set ev.Source = Me
ev.changeType = changeType
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent Change(ev)
End Sub

Private Function getChartControlIndexFromIndex(index) As Long
getChartControlIndexFromIndex = mIndexes(ChartSelector.Tabs(index).Tag)
End Function

Private Function getChartFromIndex(index) As TradeBuildChart
Set getChartFromIndex = TBChart(getChartControlIndexFromIndex(index)).Object
End Function

Private Function getIndexFromChartControlIndex(index) As Long
Dim i As Long
For i = 1 To ChartSelector.Tabs.count
    If getChartControlIndexFromIndex(i) = index Then
        getIndexFromChartControlIndex = i
        Exit For
    End If
Next
End Function

Private Sub hideChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
If index = 0 Or index > count Then Exit Sub
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then lChart.chartController.SuppressDrawing = True
TBChart(getChartControlIndexFromIndex(index)).Visible = False
End Sub

Private Sub hideTimeframeSelector()
ControlToolbar.Buttons("selecttimeframe").Width = 0
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = False
resize
End Sub

Private Function nextIndex( _
                ByVal index As Long) As Long
If index > 1 Then
    nextIndex = index - 1
ElseIf count > 1 Then
    nextIndex = 1
Else
    nextIndex = 0
End If
End Function

Private Sub resize()
If UserControl.Height < 2000 Then UserControl.Height = 2000
ControlToolbar.Left = UserControl.Width - ControlToolbar.Width
ControlToolbar.Top = UserControl.Height - ControlToolbar.Height
ChartSelector.Width = ControlToolbar.Left
ChartSelector.Top = UserControl.Height - ChartSelector.Height
ControlToolbar.ZOrder 0
ChartSelector.ZOrder 0
If count > 0 Then
    TBChart(getChartControlIndexFromIndex(ChartSelector.selectedItem.index)).Height = ChartSelector.Top
Else
    TBChart(0).Height = ChartSelector.Top
End If
End Sub

Private Sub showChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
If index = 0 Then Exit Sub
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then lChart.chartController.SuppressDrawing = False
TBChart(getChartControlIndexFromIndex(index)).Visible = True
TBChart(getChartControlIndexFromIndex(index)).Top = 0
TBChart(getChartControlIndexFromIndex(index)).Height = ChartSelector.Top
mCurrentIndex = index
End Sub

Private Sub showTimeframeSelector()
ControlToolbar.Buttons("selecttimeframe").Width = TimeframeSelector1.Width
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = True
resize
End Sub

Private Function switchToChart( _
                ByVal index As Long) As TradeBuildChart
Set switchToChart = getChartFromIndex(index)
If index <> mCurrentIndex Then
    hideChart mCurrentIndex
    showChart index
End If

fireChange MultiChartSelectionChanged

End Function

