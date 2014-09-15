VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.UserControl MultiChart 
   Alignable       =   -1  'True
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   7140
   ScaleWidth      =   9480
   Begin TradingUI27.MarketChart TBChart 
      Align           =   1  'Align Top
      Height          =   4335
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   7646
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
      Begin TradingUI27.TimeframeSelector TimeframeSelector1 
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

Event Change(ev As ChangeEventData)
Event ChartStateChanged(ByVal index As Long, ev As StateChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "MultiChart"

Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionMarketCharts                 As String = "MarketCharts"
Private Const ConfigSectionMarketChart                  As String = "MarketChart"

Private Const ConfigSettingBarFormatterFactoryName      As String = "&BarFormatterFactoryName"
Private Const ConfigSettingBarFormatterLibraryName      As String = "&BarFormatterLibraryName"
Private Const ConfigSettingChartStyle                   As String = "&ChartStyle"
Private Const ConfigSettingCurrentChart                 As String = "&CurrentChart"
Private Const ConfigSettingDataSourceKey                As String = "&DataSourceKey"

'@================================================================================
' Member variables
'@================================================================================

Private mStyle                              As ChartStyle
Private mSpec                               As ChartSpecifier
Private mIsHistoric                         As Boolean

Private mCurrentIndex                       As Long

Private mBarFormatterLibManager             As BarFormatterLibManager

Private mBarFormatterFactoryName            As String
Private mBarFormatterLibraryName            As String

Private mChangeListeners                    As Listeners

Private mConfig                             As ConfigurationSection

Private mCount                              As Long

Private mTimeframes                         As Timeframes

Private mExcludeCurrentBar                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mChangeListeners = New Listeners
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
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "MultiChart terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "MultiChart terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChartSelector_Click()
Const ProcName As String = "ChartSelector_Click"
On Error GoTo Err

switchToChart ChartSelector.SelectedItem.index
fireChange MultiChartSelectionChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ControlToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Const ProcName As String = "ControlToolbar_ButtonClick"
On Error GoTo Err

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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TBChart_StateChange(index As Integer, ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "TBChart_StateChange"
On Error GoTo Err

index = getIndexFromChartControlIndex(index)

index = getIndexFromChartControlIndex(index)

If index = mCurrentIndex And ev.State = ChartStates.ChartStateLoaded Then
    ControlToolbar.Buttons("change").Enabled = True
End If

RaiseEvent ChartStateChanged(index, ev)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TBChart_TimePeriodChange(index As Integer)
Const ProcName As String = "TBChart_TimePeriodChange"
On Error GoTo Err

ChartSelector.Tabs(getIndexFromChartControlIndex(index)).caption = TBChart(index).TimePeriod.ToShortString
fireChange MultiChartPeriodLengthChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TimeframeSelector1_Click()
Const ProcName As String = "TimeframeSelector1_Click"
On Error GoTo Err

If ControlToolbar.Buttons("add").value = tbrPressed Then
    Add TimeframeSelector1.TimePeriod
    hideTimeframeSelector
    ControlToolbar.Buttons("add").value = tbrUnpressed
Else
    Chart.ChangeTimePeriod TimeframeSelector1.TimePeriod
    hideTimeframeSelector
    ControlToolbar.Buttons("change").value = tbrUnpressed
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

Public Property Get BaseChartController( _
                Optional ByVal index As Long = -1) As ChartController
Const ProcName As String = "BaseChartController"
On Error GoTo Err

index = checkIndex(index)
Set BaseChartController = getChartFromIndex(index).BaseChartController

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' do not make this Public because the value returned cannot be handled by non-friend
' components
Friend Property Get Chart( _
                Optional ByVal index As Long = -1) As MarketChart
Const ProcName As String = "Chart"
On Error GoTo Err

index = checkIndex(index)
Set Chart = getChartFromIndex(index)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartManager( _
                Optional ByVal index As Long = -1) As ChartManager
Const ProcName As String = "ChartManager"
On Error GoTo Err

index = checkIndex(index)
Set ChartManager = getChartFromIndex(index).ChartManager

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mCount

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If mConfig Is value Then Exit Property
If Not mConfig Is Nothing Then mConfig.Remove
If value Is Nothing Then Exit Property

Set mConfig = value

gLogger.Log "MultiChart added to config at: " & mConfig.Path, ProcName, ModuleName

storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LoadingText( _
                Optional ByVal index As Long = -1) As Text
Const ProcName As String = "LoadingText"
On Error GoTo Err

index = checkIndex(index)
Set LoadingText = getChartFromIndex(index).LoadingText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PriceRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Const ProcName As String = "PriceRegion"
On Error GoTo Err

index = checkIndex(index)
Set PriceRegion = getChartFromIndex(index).PriceRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Style(ByVal pStyle As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Dim i As Long
Dim lChart As MarketChart

Set mStyle = pStyle

If Not mConfig Is Nothing Then
    If mStyle Is Nothing Then
        mConfig.SetSetting ConfigSettingChartStyle, ""
    Else
        mConfig.SetSetting ConfigSettingChartStyle, mStyle.Name
    End If
End If

For i = 1 To TBChart.UBound
    Set lChart = TBChart(i).object
    lChart.BaseChartController.Style = mStyle
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get State( _
                Optional ByVal index As Long = -1) As ChartStates
Const ProcName As String = "State"
On Error GoTo Err

index = checkIndex(index)
State = getChartFromIndex(index).State

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Timeframe( _
                Optional ByVal index As Long = -1) As Timeframe
Const ProcName As String = "Timeframe"
On Error GoTo Err

index = checkIndex(index)
Set Timeframe = getChartFromIndex(index).Timeframe

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod( _
                Optional ByVal index As Long = -1) As TimePeriod
Const ProcName As String = "TimePeriod"
On Error GoTo Err

index = checkIndex(index)
Set TimePeriod = getChartFromIndex(index).TimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Const ProcName As String = "VolumeRegion"
On Error GoTo Err

index = checkIndex(index)
Set VolumeRegion = getChartFromIndex(index).VolumeRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

' return the index rather than the chart because the chart cannot be handled by non-friend
' components
Public Function Add( _
                ByVal pPeriodLength As TimePeriod, _
                Optional ByVal pTitle As String, _
                Optional ByVal pUpdatePerTick As Boolean = True) As Long
Const ProcName As String = "Add"
On Error GoTo Err

Dim lChart As MarketChart
Dim lTab As MSComctlLib.Tab

Set lChart = loadChartControl
lChart.UpdatePerTick = pUpdatePerTick

Set lTab = addTab(pPeriodLength)

' we notify the add before calling ShowChart so that it's before
' the ChartStates.ChartStateInitialised and ChartStates.ChartStateLoaded events
fireChange MultiChartAdd

lChart.ShowChart mTimeframes, pPeriodLength, mSpec, mStyle, mBarFormatterLibManager, mBarFormatterFactoryName, mBarFormatterLibraryName, mExcludeCurrentBar, pTitle

If Not mConfig Is Nothing Then
    lChart.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionMarketCharts).AddConfigurationSection(ConfigSectionMarketChart & "(" & GenerateGUIDString & ")")
End If

lTab.Selected = True
Add = mCurrentIndex

fireChange MultiChartSelectionChanged

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddChangeListener( _
                ByVal pListener As ChangeListener)
Const ProcName As String = "AddChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
               
Public Sub ChangePeriodLength(ByVal pNewTimePeriod As TimePeriod)
Const ProcName As String = "ChangePeriodLength"
On Error GoTo Err

Chart.ChangeTimePeriod pNewTimePeriod

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Do While ChartSelector.Tabs.Count <> 0
    Remove ChartSelector.Tabs.Count
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FocusChart()
Const ProcName As String = "FocusChart"
On Error GoTo Err

Dim index As Long
index = checkIndex(-1)
TBChart(getChartControlIndexFromIndex(index)).SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim i As Long
For i = 1 To mCount
    getChartFromIndex(i).Finish
    unloadChartControl i
Next
TBChart(0).Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pTimeframes As Timeframes, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pSpec As ChartSpecifier, _
                Optional ByVal pStyle As ChartStyle, _
                Optional ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String, _
                Optional ByVal pExcludeCurrentBar As Boolean, _
                Optional ByVal pBackColor As Long = &HC0C0C0)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"

Set mTimeframes = pTimeframes

Set mSpec = pSpec

If pStyle Is Nothing Then
    Set mStyle = ChartStylesManager.DefaultStyle
Else
    Set mStyle = pStyle
End If

TBChart(0).ChartBackColor = pBackColor
mIsHistoric = (mSpec.toTime <> 0)

Set mBarFormatterLibManager = pBarFormatterLibManager
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName

mExcludeCurrentBar = pExcludeCurrentBar

TimeframeSelector1.Initialise pTimePeriodValidator

storeSettings

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function LoadFromConfig( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pTimeframes As Timeframes, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                Optional ByVal pBarFormatterLibManager As BarFormatterLibManager) As Boolean
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

AssertArgument Not pConfig Is Nothing, "pConfig cannot be Nothing"

Set mConfig = pConfig

gLogger.Log "Loading MultiChart from config at: " & mConfig.Path, ProcName, ModuleName

Set mTimeframes = pTimeframes
Set mBarFormatterLibManager = pBarFormatterLibManager
TimeframeSelector1.Initialise pTimePeriodValidator

Set mSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))

Dim lStyleName As String
lStyleName = mConfig.GetSetting(ConfigSettingChartStyle, "")
If ChartStylesManager.Contains(lStyleName) Then
    Set mStyle = ChartStylesManager.Item(lStyleName)
Else
    Set mStyle = ChartStylesManager.DefaultStyle
End If

mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")

mIsHistoric = (mSpec.toTime <> 0)

Dim currentChartIndex As Long
currentChartIndex = CLng(mConfig.GetSetting(ConfigSettingCurrentChart, "1"))

Dim cs As ConfigurationSection
For Each cs In mConfig.AddConfigurationSection(ConfigSectionMarketCharts)
    AddFromConfig cs
Next

If ChartSelector.Tabs.Count > 0 Then SelectChart currentChartIndex

LoadFromConfig = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pModuleName:=ModuleName
LoadFromConfig = False
End Function

Public Sub Remove( _
                ByVal index As Long)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim nxtIndex As Long

AssertArgument index <= Count And index >= 1, "Index must not be less than 1 or greater than Count"

nxtIndex = nextIndex(index)
closeChart index
If index = mCurrentIndex Then mCurrentIndex = 0

If nxtIndex <> 0 Then
    SelectChart nxtIndex
Else
    If Not mConfig Is Nothing Then mConfig.RemoveSetting ConfigSettingCurrentChart
End If

fireChange MultiChartRemove
fireChange MultiChartSelectionChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal pListener As ChangeListener)
Const ProcName As String = "RemoveChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

gLogger.Log "MultiChart removed from config at: " & mConfig.Path, ProcName, ModuleName

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"
On Error GoTo Err

Chart.ScrollToTime pTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectChart( _
                ByVal index As Long)
Const ProcName As String = "SelectChart"
On Error GoTo Err

AssertArgument index <= Count And index >= 1, "Index must not be less than 1 or greater than Count"

ChartSelector.Tabs(index).Selected = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub AddFromConfig( _
                ByVal chartSect As ConfigurationSection)
Const ProcName As String = "AddFromConfig"
On Error GoTo Err

Dim lChart As MarketChart
Set lChart = loadChartControl

Dim lTab As MSComctlLib.Tab
Set lTab = addTab(Nothing)

lChart.LoadFromConfig mTimeframes, chartSect, mBarFormatterLibManager, True

lTab.caption = lChart.TimePeriod.ToShortString

fireChange MultiChartAdd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function addTab( _
                ByVal pPeriodLength As TimePeriod) As MSComctlLib.Tab

Const ProcName As String = "addTab"
On Error GoTo Err

If pPeriodLength Is Nothing Then
    Set addTab = ChartSelector.Tabs.Add(, , "")
Else
    Set addTab = ChartSelector.Tabs.Add(, , pPeriodLength.ToShortString)
End If
addTab.Tag = TBChart.UBound

ControlToolbar.Buttons("remove").Enabled = True

If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCurrentChart, ChartSelector.Tabs.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function checkIndex( _
                ByVal index As Long) As Long
Const ProcName As String = "checkIndex"
On Error GoTo Err

If index = -1 Then
    Assert mCurrentIndex >= 1, "No current chart"
    index = mCurrentIndex
End If

AssertArgument index <= Count And index >= 1, "Index must not be less than 1 or greater than Count"

checkIndex = index

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub closeChart( _
                ByVal index As Long)
Const ProcName As String = "closeChart"
On Error GoTo Err

Dim lChart As MarketChart
Set lChart = getChartFromIndex(index)
lChart.RemoveFromConfig
lChart.Finish
unloadChartControl index
ChartSelector.Tabs.Remove index
If ChartSelector.Tabs.Count = 0 Then
    ControlToolbar.Buttons("remove").Enabled = False
    TBChart(0).Visible = True
    TBChart(0).Top = 0
    TBChart(0).Height = ChartSelector.Top
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal changeType As MultiChartChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.changeType = changeType

mChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As ChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getChartControlIndexFromIndex(index) As Long
Const ProcName As String = "getChartControlIndexFromIndex"
On Error GoTo Err

getChartControlIndexFromIndex = ChartSelector.Tabs(index).Tag

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getChartFromIndex(index) As MarketChart
Const ProcName As String = "getChartFromIndex"
On Error GoTo Err

Set getChartFromIndex = TBChart(getChartControlIndexFromIndex(index)).object

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getIndexFromChartControlIndex(index) As Long
Const ProcName As String = "getIndexFromChartControlIndex"
On Error GoTo Err

Dim i As Long
For i = 1 To ChartSelector.Tabs.Count
    If getChartControlIndexFromIndex(i) = index Then
        getIndexFromChartControlIndex = i
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideChart( _
                ByVal index As Long)
Const ProcName As String = "hideChart"
On Error GoTo Err

Dim lChart As MarketChart

If index = 0 Or index > Count Then Exit Sub
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then lChart.DisableDrawing
TBChart(getChartControlIndexFromIndex(index)).Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub hideTimeframeSelector()
Const ProcName As String = "hideTimeframeSelector"
On Error GoTo Err

ControlToolbar.Buttons("selecttimeframe").Width = 0
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = False
resize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function loadChartControl() As MarketChart
Const ProcName As String = "loadChartControl"
On Error GoTo Err

Load TBChart(TBChart.UBound + 1)
TBChart(TBChart.UBound).align = vbAlignTop
TBChart(TBChart.UBound).Top = 0
TBChart(TBChart.UBound).Height = ChartSelector.Top
mCount = mCount + 1
Set loadChartControl = TBChart(TBChart.UBound).object

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function nextIndex( _
                ByVal index As Long) As Long
Const ProcName As String = "nextIndex"
On Error GoTo Err

If index > 1 Then
    nextIndex = index - 1
ElseIf Count > 1 Then
    nextIndex = 1
Else
    nextIndex = 0
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ShowChart( _
                ByVal index As Long) As MarketChart
Const ProcName As String = "ShowChart"
On Error GoTo Err

Dim lChart As MarketChart

If index = 0 Then Exit Function

Set lChart = getChartFromIndex(index)

If lChart.State = ChartStates.ChartStateBlank Then lChart.Start

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

TimeframeSelector1.SelectTimeframe lChart.TimePeriod

Set ShowChart = lChart

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showTimeframeSelector()
Const ProcName As String = "showTimeframeSelector"
On Error GoTo Err

ControlToolbar.Buttons("selecttimeframe").Width = TimeframeSelector1.Width
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = True
resize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)

If Not mStyle Is Nothing Then mConfig.SetSetting ConfigSettingChartStyle, mStyle.Name

mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName

Dim cs As ConfigurationSection
Set cs = mConfig.AddConfigurationSection(ConfigSectionMarketCharts)

Dim i As Long
For i = 1 To TBChart.UBound
    If Not TBChart(i) Is Nothing Then
        Dim lChart As MarketChart
        Set lChart = TBChart(i).object
        lChart.ConfigurationSection = cs.AddConfigurationSection(ConfigSectionMarketChart & "(" & GenerateGUIDString & ")")
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub switchToChart( _
                ByVal index As Long)
Const ProcName As String = "switchToChart"
On Error GoTo Err

If index = mCurrentIndex Then Exit Sub

hideChart mCurrentIndex
ShowChart index

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unloadChartControl(ByVal index As Long)
Const ProcName As String = "unloadChartControl"
On Error GoTo Err

Unload TBChart(getChartControlIndexFromIndex(index))
mCount = mCount - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
