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
      Height          =   4095
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   7223
      ChartBackColor  =   6566450
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
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "selecttimeframe"
            Object.ToolTipText     =   "Choose the timeframe for the new chart"
            Style           =   4
            Object.Width           =   1700
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "change"
            Object.ToolTipText     =   "Change the timeframe for the current chart"
            ImageIndex      =   3
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Select a new timeframe and add another chart"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Object.ToolTipText     =   "Remove current chart"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin TradeBuildUI26.TimeframeSelector TimeframeSelector1 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
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
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":08A4
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
Event ChartStateChanged(ByVal index As Long, ev As StateChangeEvent)

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

Private Const ConfigSectionBarFormatterFactory          As String = "BarFormatterFactory"
Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionTradeBuildCharts             As String = "TradeBuildCharts"
Private Const ConfigSectionTradeBuildChart              As String = "TradeBuildChart"

Private Const ConfigSettingCurrentChart                 As String = ".CurrentChart"
Private Const ConfigSettingFromTime                     As String = ".FromTime"
Private Const ConfigSettingProgId                       As String = "&ProgId"
Private Const ConfigSettingToTime                       As String = ".ToTime"
Private Const ConfigSettingTickerKey                    As String = ".TickerKey"
Private Const ConfigSettingWorkspace                    As String = ".Workspace"

'@================================================================================
' Member variables
'@================================================================================

Private mTicker                             As Ticker
Private mSpec                               As ChartSpecifier
Private mIsHistoric                         As Boolean
Private mFromTime                           As Date
Private mToTime                             As Date

Private mIndexes                            As Collection
Private mCurrentIndex                       As Long

Private mBarFormatterFactory                As BarFormatterFactory

Private mChangeListeners                    As Collection

Private mConfig                             As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mIndexes = New Collection
Set mChangeListeners = New Collection
ChartSelector.Tabs.Clear
TBChart(0).Visible = True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
hideTimeframeSelector
End Sub

Private Sub UserControl_Resize()
resize
End Sub

Private Sub UserControl_Terminate()
gLogger.Log LogLevelDetail, "MultiChart terminated"
Debug.Print "MultiChart terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChartSelector_Click()
switchToChart ChartSelector.SelectedItem.index
fireChange MultiChartSelectionChanged
End Sub

Private Sub ControlToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case UCase$(Button.Key)
Case "ADD"
    If ControlToolbar.Buttons("add").value = tbrPressed Then
        ControlToolbar.Buttons("change").value = tbrUnpressed
        showTimeframeSelector
    Else
        hideTimeframeSelector
    End If
Case "CHANGE"
    If ControlToolbar.Buttons("change").value = tbrPressed Then
        ControlToolbar.Buttons("add").value = tbrUnpressed
        showTimeframeSelector
    Else
        hideTimeframeSelector
    End If
Case "REMOVE"
    ControlToolbar.Buttons("add").value = tbrUnpressed
    ControlToolbar.Buttons("change").value = tbrUnpressed
    Remove mCurrentIndex
End Select
End Sub

Private Sub TBChart_StateChange(index As Integer, ev As TWUtilities30.StateChangeEvent)

If index = mCurrentIndex And ev.State = ChartStates.ChartStateLoaded Then
    ControlToolbar.Buttons("change").Enabled = True
End If

RaiseEvent ChartStateChanged(index, ev)
End Sub

Private Sub TBChart_TimeframeChange(index As Integer)
ChartSelector.Tabs(getIndexFromChartControlIndex(index)).caption = TBChart(index).TimePeriod.ToShortString
fireChange MultiChartTimeframeChanged
End Sub

Private Sub TimeframeSelector1_Click()
If ControlToolbar.Buttons("add").value = tbrPressed Then
    Add TimeframeSelector1.timeframeDesignator
    hideTimeframeSelector
    ControlToolbar.Buttons("add").value = tbrUnpressed
Else
    Chart.ChangeTimeframe TimeframeSelector1.timeframeDesignator
    hideTimeframeSelector
    ControlToolbar.Buttons("change").value = tbrUnpressed
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get BaseChartController( _
                Optional ByVal index As Long = -1) As ChartController
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set BaseChartController = getChartFromIndex(index).BaseChartController

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "BaseChartController" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

' do not make this Public because the value returned cannot be handled by non-friend
' components
Friend Property Get Chart( _
                Optional ByVal index As Long = -1) As TradeBuildChart
index = checkIndex(index)
Set Chart = getChartFromIndex(index)
End Property

Public Property Get ChartManager( _
                Optional ByVal index As Long = -1) As ChartManager
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set ChartManager = getChartFromIndex(index).ChartManager

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "chartManager" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Count() As Long
Count = ChartSelector.Tabs.Count
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings
End Property

Public Property Get Enabled() As Boolean
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
UserControl.Enabled = value
PropertyChanged "Enabled"
End Property

Public Property Get LoadingText( _
                Optional ByVal index As Long = -1) As Text
index = checkIndex(index)
Set LoadingText = getChartFromIndex(index).LoadingText
End Property

Public Property Get PriceRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set PriceRegion = getChartFromIndex(index).PriceRegion

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "priceRegion" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "state" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get Ticker() As Ticker
Set Ticker = mTicker
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Timeframe" & "." & failpoint
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
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "TimePeriod" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

Public Property Get VolumeRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set VolumeRegion = getChartFromIndex(index).VolumeRegion

Exit Property

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "VolumeRegion" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pTimeframe As TimePeriod) As TradeBuildChart
Dim lChart As TradeBuildChart
Dim lSpec As ChartSpecifier
Dim lTab As MSComctlLib.Tab
Dim failpoint As Long
On Error GoTo Err

load TBChart(TBChart.UBound + 1)
Set lChart = TBChart(TBChart.UBound).object
TBChart(TBChart.UBound).align = vbAlignTop
TBChart(TBChart.UBound).Top = 0
TBChart(TBChart.UBound).Height = ChartSelector.Top
Set lSpec = mSpec.Clone
lSpec.Timeframe = pTimeframe

Set lTab = addTab(pTimeframe)

' we notify the add before calling ShowChart or ShowHistoric chart so that it's before
' the ChartStates.ChartStateInitialised and ChartStates.ChartStateLoaded events
fireChange MultiChartAdd

If Not mConfig Is Nothing Then
    lChart.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts).AddConfigurationSection(ConfigSectionTradeBuildChart & "(" & GenerateGUIDString & ")")
End If

If mIsHistoric Then
    lChart.showHistoricChart mTicker, lSpec, mFromTime, mToTime, mBarFormatterFactory
Else
    lChart.showChart mTicker, lSpec, mBarFormatterFactory
End If

lTab.Selected = True
fireChange MultiChartSelectionChanged

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Add" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Function

Public Sub AddChangeListener( _
                ByVal listener As ChangeListener)
mChangeListeners.Add listener, CStr(ObjPtr(listener))
End Sub
               
Public Sub ChangeTimeframe(ByVal pTimeframe As TimePeriod)
Dim failpoint As Long
On Error GoTo Err

Chart.ChangeTimeframe pTimeframe

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ChangeTimeframe" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub Clear()
Do While ChartSelector.Tabs.Count <> 0
    Remove ChartSelector.Tabs.Count
Loop
End Sub

Public Sub Finish()
Dim i As Long
Dim index As Long
For i = 1 To mIndexes.Count
    index = mIndexes(i)
    getChartFromIndex(index).Finish
    Unload TBChart(getChartControlIndexFromIndex(index))
Next
TBChart(0).Finish
End Sub

Public Sub Initialise( _
                ByVal pTicker As Ticker, _
                ByVal chartSpec As ChartSpecifier, _
                Optional ByVal fromTime As Date, _
                Optional ByVal toTime As Date, _
                Optional ByVal BarFormatterFactory As BarFormatterFactory)
Dim failpoint As Long
On Error GoTo Err

Set mTicker = pTicker
Set mSpec = chartSpec
TBChart(0).ChartBackColor = mSpec.ChartBackColor
mIsHistoric = (fromTime <> 0 Or toTime <> 0)
mFromTime = fromTime
mToTime = toTime
Set mBarFormatterFactory = BarFormatterFactory

TimeframeSelector1.Initialise
TimeframeSelector1.selectTimeframe mSpec.Timeframe

storeSettings

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Initialise" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Function LoadFromConfig( _
                ByVal config As ConfigurationSection) As Boolean
Dim cs As ConfigurationSection
Dim currentChartIndex As Long

Dim failpoint As Long
On Error GoTo Err

Set mConfig = config
If mConfig.GetSetting(ConfigSettingWorkspace) = "" Or _
    mConfig.GetSetting(ConfigSettingTickerKey) = "" _
Then
    Exit Function
End If

Set mTicker = TradeBuildAPI.WorkSpaces(mConfig.GetSetting(ConfigSettingWorkspace)).Tickers(mConfig.GetSetting(ConfigSettingTickerKey))
Set mSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))
mFromTime = CDate(mConfig.GetSetting(ConfigSettingFromTime, "0"))
mToTime = CDate(mConfig.GetSetting(ConfigSettingToTime, "0"))

mIsHistoric = (mFromTime <> 0 Or mToTime <> 0)

Set cs = mConfig.GetConfigurationSection(ConfigSectionBarFormatterFactory)
If Not cs Is Nothing Then
    Set mBarFormatterFactory = CreateObject(cs.GetSetting(ConfigSettingProgId))
    mBarFormatterFactory.LoadFromConfig cs
End If

TimeframeSelector1.Initialise
TimeframeSelector1.selectTimeframe mSpec.Timeframe

currentChartIndex = CLng(mConfig.GetSetting(ConfigSettingCurrentChart, "1"))

For Each cs In mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts)
    addFromConfig cs
Next

If ChartSelector.Tabs.Count > 0 Then SelectChart currentChartIndex

LoadFromConfig = True

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "LoadFromConfig" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

LoadFromConfig = False
End Function

Public Sub Remove( _
                ByVal index As Long)
Dim nxtIndex As Long
Dim failpoint As Long
On Error GoTo Err

If index > Count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "Remove", _
            "Index must not be less than 1 or greater than Count"
End If

nxtIndex = nextIndex(index)
closeChart index
If index = mCurrentIndex Then mCurrentIndex = 0

If nxtIndex <> 0 Then
    SelectChart nxtIndex
Else
    If Not mConfig Is Nothing Then mConfig.RemoveSetting (ConfigSettingCurrentChart)
End If

fireChange MultiChartRemove
fireChange MultiChartSelectionChanged

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "Remove" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub RemoveChangeListener(ByVal listener As ChangeListener)
mChangeListeners.Remove ObjPtr(listener)
End Sub

Public Sub RemoveFromConfig()
mConfig.Remove
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Dim failpoint As Long
On Error GoTo Err

Chart.ScrollToTime pTime

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "ScrollToTime" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SelectChart( _
                ByVal index As Long)
Dim failpoint As Long
On Error GoTo Err

If index > Count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SelectChart", _
            "Index must not be less than 1 or greater than Count"
End If

ChartSelector.Tabs(index).Selected = True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = IIf(Err.Source <> "", Err.Source & vbCrLf, "") & ProjectName & "." & ModuleName & ":" & "SelectChart" & "." & failpoint
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addFromConfig( _
                ByVal chartSect As ConfigurationSection)
Dim lChart As TradeBuildChart
Dim failpoint As Long

load TBChart(TBChart.UBound + 1)
Set lChart = TBChart(TBChart.UBound).object
TBChart(TBChart.UBound).align = vbAlignTop
TBChart(TBChart.UBound).Top = 0
TBChart(TBChart.UBound).Height = ChartSelector.Top

lChart.LoadFromConfig chartSect

addTab(lChart.TimePeriod).Selected = True
End Sub

Private Function addTab( _
                ByVal pTimePeriod As TimePeriod) As MSComctlLib.Tab
Static chartNumber As Long

chartNumber = chartNumber + 1
Set addTab = ChartSelector.Tabs.Add(, , pTimePeriod.ToShortString)
addTab.Tag = CStr(chartNumber)

ControlToolbar.Buttons("remove").Enabled = True

mIndexes.Add TBChart.UBound, CStr(chartNumber)
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCurrentChart, ChartSelector.Tabs.Count
End Function

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

If index > Count Or index < 1 Then
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
lChart.RemoveFromConfig
lChart.Finish
Unload TBChart(getChartControlIndexFromIndex(index))
mIndexes.Remove ChartSelector.Tabs(index).Tag
ChartSelector.Tabs.Remove index
If ChartSelector.Tabs.Count = 0 Then
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
Set getChartFromIndex = TBChart(getChartControlIndexFromIndex(index))
End Function

Private Function getIndexFromChartControlIndex(index) As Long
Dim i As Long
For i = 1 To ChartSelector.Tabs.Count
    If getChartControlIndexFromIndex(i) = index Then
        getIndexFromChartControlIndex = i
        Exit For
    End If
Next
End Function

Private Sub hideChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
If index = 0 Or index > Count Then Exit Sub
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then lChart.DisableDrawing
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
ElseIf Count > 1 Then
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
If Count > 0 Then
    TBChart(getChartControlIndexFromIndex(ChartSelector.SelectedItem.index)).Height = ChartSelector.Top
Else
    TBChart(0).Height = ChartSelector.Top
End If
End Sub

Private Function showChart( _
                ByVal index As Long) As TradeBuildChart
Dim lChart As TradeBuildChart
If index = 0 Then Exit Function
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then
    lChart.EnableDrawing
    ControlToolbar.Buttons("change").Enabled = True
Else
    ControlToolbar.Buttons("change").Enabled = False
End If

TBChart(getChartControlIndexFromIndex(index)).Visible = True
TBChart(getChartControlIndexFromIndex(index)).Top = 0
TBChart(getChartControlIndexFromIndex(index)).Height = ChartSelector.Top
mCurrentIndex = index

If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCurrentChart, index

Set showChart = lChart
End Function

Private Sub showTimeframeSelector()
ControlToolbar.Buttons("selecttimeframe").Width = TimeframeSelector1.Width
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = True
resize
End Sub

Private Sub storeSettings()
Dim i As Long
Dim lChart As TradeBuildChart
Dim cs As ConfigurationSection

If mConfig Is Nothing Then Exit Sub

If Not mTicker Is Nothing Then
    mConfig.SetSetting ConfigSettingWorkspace, mTicker.Workspace.name
    mConfig.SetSetting ConfigSettingTickerKey, mTicker.Key
End If

mConfig.SetSetting ConfigSettingFromTime, CStr(CDbl(mFromTime))
mConfig.SetSetting ConfigSettingToTime, CStr(CDbl(mToTime))

If Not mBarFormatterFactory Is Nothing Then
    Set cs = mConfig.AddConfigurationSection(ConfigSectionBarFormatterFactory)
    cs.SetSetting ConfigSettingProgId, GetProgIdFromObject(mBarFormatterFactory)
    mBarFormatterFactory.ConfigurationSection = cs
End If

If Not mSpec Is Nothing Then mSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)

Set cs = mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts)
For i = 1 To TBChart.UBound
    If Not TBChart(i) Is Nothing Then
        Set lChart = TBChart(i).object
        lChart.ConfigurationSection = cs.AddConfigurationSection(ConfigSectionTradeBuildChart & "(" & GenerateGUIDString & ")")
    End If
Next
End Sub

Private Sub switchToChart( _
                ByVal index As Long)
If index = mCurrentIndex Then Exit Sub

hideChart mCurrentIndex
showChart index

End Sub

