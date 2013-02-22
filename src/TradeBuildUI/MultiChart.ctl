VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.UserControl MultiChart 
   Alignable       =   -1  'True
   ClientHeight    =   7140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   7140
   ScaleWidth      =   9480
   Begin TradeBuildUI27.TradeBuildChart TBChart 
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
      ChartBackColor  =   0
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
      Begin TradeBuildUI27.TimeframeSelector TimeframeSelector1 
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

Private Const ModuleName                    As String = "MultiChart"

Private Const ConfigSectionChartSpecifier               As String = "ChartSpecifier"
Private Const ConfigSectionTradeBuildCharts             As String = "TradeBuildCharts"
Private Const ConfigSectionTradeBuildChart              As String = "TradeBuildChart"

Private Const ConfigSettingBarFormatterFactoryName      As String = "&BarFormatterFactoryName"
Private Const ConfigSettingBarFormatterLibraryName      As String = "&BarFormatterLibraryName"
Private Const ConfigSettingChartStyle                   As String = "&ChartStyle"
Private Const ConfigSettingCurrentChart                 As String = "&CurrentChart"
Private Const ConfigSettingTickerKey                    As String = "&TickerKey"
Private Const ConfigSettingWorkspace                    As String = "&Workspace"

'@================================================================================
' Member variables
'@================================================================================

Private mTicker                             As Ticker
Private mStyle                              As ChartStyle
Private mSpec                               As ChartSpecifier
Private mIsHistoric                         As Boolean

Private mCurrentIndex                       As Long

Private mBarFormatterFactoryName            As String
Private mBarFormatterLibraryName            As String

Private mChangeListeners                    As Collection

Private mConfig                             As ConfigurationSection

Private mCount                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
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
Dim failpoint As Long
On Error GoTo Err

switchToChart ChartSelector.SelectedItem.index
fireChange MultiChartSelectionChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub ControlToolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Const ProcName As String = "ControlToolbar_ButtonClick"
Dim failpoint As Long
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

Private Sub TBChart_StateChange(index As Integer, ev As TWUtilities30.StateChangeEventData)
Const ProcName As String = "TBChart_StateChange"
Dim failpoint As Long
On Error GoTo Err

index = getIndexFromChartControlIndex(index)

If index = mCurrentIndex And ev.State = ChartStates.ChartStateLoaded Then
    ControlToolbar.Buttons("change").Enabled = True
End If

RaiseEvent ChartStateChanged(index, ev)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TBChart_TimeframeChange(index As Integer)
Const ProcName As String = "TBChart_TimeframeChange"
Dim failpoint As Long
On Error GoTo Err

ChartSelector.Tabs(getIndexFromChartControlIndex(index)).caption = TBChart(index).PeriodLength.ToShortString
fireChange MultiChartPeriodLengthChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TimeframeSelector1_Click()
Const ProcName As String = "TimeframeSelector1_Click"
Dim failpoint As Long
On Error GoTo Err

If ControlToolbar.Buttons("add").value = tbrPressed Then
    Add TimeframeSelector1.TimeframeDesignator
    hideTimeframeSelector
    ControlToolbar.Buttons("add").value = tbrUnpressed
Else
    Chart.ChangePeriodLength TimeframeSelector1.TimeframeDesignator
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
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set BaseChartController = getChartFromIndex(index).BaseChartController

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

' do not make this Public because the value returned cannot be handled by non-friend
' components
Friend Property Get Chart( _
                Optional ByVal index As Long = -1) As TradeBuildChart
Const ProcName As String = "Chart"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set Chart = getChartFromIndex(index)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get ChartManager( _
                Optional ByVal index As Long = -1) As ChartManager
Const ProcName As String = "ChartManager"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set ChartManager = getChartFromIndex(index).ChartManager

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
Dim failpoint As Long
On Error GoTo Err

Count = mCount

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As Long
On Error GoTo Err

If value Is mConfig Then Exit Property
Set mConfig = value
storeSettings

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Enabled() As Boolean
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Enabled( _
                ByVal value As Boolean)
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

UserControl.Enabled = value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LoadingText( _
                Optional ByVal index As Long = -1) As Text
Const ProcName As String = "LoadingText"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set LoadingText = getChartFromIndex(index).LoadingText

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PriceRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Const ProcName As String = "PriceRegion"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set PriceRegion = getChartFromIndex(index).PriceRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Style(ByVal pStyle As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Dim i As Long
Dim lChart As TradeBuildChart

Set mStyle = pStyle

If Not mConfig Is Nothing Then
    If mStyle Is Nothing Then
        mConfig.SetSetting ConfigSettingChartStyle, ""
    Else
        mConfig.SetSetting ConfigSettingChartStyle, mStyle.name
    End If
End If

For i = 1 To TBChart.UBound
    Set lChart = TBChart(i).object
    lChart.BaseChartController.Style = mStyle
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get State( _
                Optional ByVal index As Long = -1) As ChartStates
Const ProcName As String = "State"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
State = getChartFromIndex(index).State

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Ticker() As Ticker
Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set Ticker = mTicker

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Timeframe( _
                Optional ByVal index As Long = -1) As Timeframe
Const ProcName As String = "Timeframe"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set Timeframe = getChartFromIndex(index).Timeframe

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get TimePeriod( _
                Optional ByVal index As Long = -1) As TimePeriod
Const ProcName As String = "TimePeriod"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set TimePeriod = getChartFromIndex(index).PeriodLength

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get VolumeRegion( _
                Optional ByVal index As Long = -1) As ChartRegion
Const ProcName As String = "VolumeRegion"
Dim failpoint As Long
On Error GoTo Err

index = checkIndex(index)
Set VolumeRegion = getChartFromIndex(index).VolumeRegion

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add( _
                ByVal pPeriodLength As TimePeriod) As TradeBuildChart
Dim lChart As TradeBuildChart
Dim lTab As MSComctlLib.Tab

Const ProcName As String = "Add"
Dim failpoint As Long
On Error GoTo Err

Set lChart = loadChartControl

Set lTab = addTab(pPeriodLength)

' we notify the add before calling ShowChart so that it's before
' the ChartStates.ChartStateInitialised and ChartStates.ChartStateLoaded events
fireChange MultiChartAdd

If Not mConfig Is Nothing Then
    lChart.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts).AddConfigurationSection(ConfigSectionTradeBuildChart & "(" & GenerateGUIDString & ")")
End If

lChart.ShowChart mTicker, pPeriodLength, mSpec, mStyle, mBarFormatterFactoryName, mBarFormatterLibraryName

lTab.Selected = True

fireChange MultiChartSelectionChanged

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub AddChangeListener( _
                ByVal listener As ChangeListener)
Const ProcName As String = "AddChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Add listener, CStr(ObjPtr(listener))

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
               
Public Sub ChangePeriodLength(ByVal pNewPeriodLength As TimePeriod)
Const ProcName As String = "ChangePeriodLength"
Dim failpoint As Long
On Error GoTo Err

Chart.ChangePeriodLength pNewPeriodLength

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
Dim failpoint As Long
On Error GoTo Err

Do While ChartSelector.Tabs.Count <> 0
    Remove ChartSelector.Tabs.Count
Loop

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Finish()
Dim i As Long
Const ProcName As String = "Finish"
On Error GoTo Err

For i = 1 To mCount
    getChartFromIndex(i).Finish
    unloadChartControl i
Next
TBChart(0).Finish

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Initialise( _
                ByVal pTicker As Ticker, _
                Optional ByVal pSpec As ChartSpecifier, _
                Optional ByVal pStyle As ChartStyle, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String, _
                Optional ByVal pBackColor As Long = &HC0C0C0)
Const ProcName As String = "Initialise"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = pTicker

Set mSpec = pSpec

If pStyle Is Nothing Then
    Set mStyle = ChartStylesManager.DefaultStyle
Else
    Set mStyle = pStyle
End If

TBChart(0).ChartBackColor = pBackColor
mIsHistoric = (mSpec.toTime <> 0)
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName

TimeframeSelector1.Initialise

storeSettings

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function LoadFromConfig( _
                ByVal config As ConfigurationSection) As Boolean
Dim cs As ConfigurationSection
Dim currentChartIndex As Long

Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig.GetSetting(ConfigSettingWorkspace) = "" Or _
    mConfig.GetSetting(ConfigSettingTickerKey) = "" _
Then
    Exit Function
End If

Set mTicker = TradeBuildAPI.WorkSpaces(mConfig.GetSetting(ConfigSettingWorkspace)).Tickers(mConfig.GetSetting(ConfigSettingTickerKey))
Set mSpec = LoadChartSpecifierFromConfig(mConfig.GetConfigurationSection(ConfigSectionChartSpecifier))

Dim lStyleName As String
lStyleName = mConfig.GetSetting(ConfigSettingChartStyle, "")
If ChartStylesManager.Contains(lStyleName) Then
    Set mStyle = ChartStylesManager.item(lStyleName)
Else
    Set mStyle = ChartStylesManager.DefaultStyle
End If

mBarFormatterFactoryName = mConfig.GetSetting(ConfigSettingBarFormatterFactoryName, "")
mBarFormatterLibraryName = mConfig.GetSetting(ConfigSettingBarFormatterLibraryName, "")

mIsHistoric = (mSpec.toTime <> 0)

TimeframeSelector1.Initialise

currentChartIndex = CLng(mConfig.GetSetting(ConfigSettingCurrentChart, "1"))

For Each cs In mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts)
    addFromConfig cs
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
Dim nxtIndex As Long

Const ProcName As String = "Remove"
Dim failpoint As Long
On Error GoTo Err

If index > Count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Index must not be less than 1 or greater than Count"
End If

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveChangeListener(ByVal listener As ChangeListener)
Const ProcName As String = "RemoveChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Remove ObjPtr(listener)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As Long
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ScrollToTime(ByVal pTime As Date)
Const ProcName As String = "ScrollToTime"
Dim failpoint As Long
On Error GoTo Err

Chart.ScrollToTime pTime

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub SelectChart( _
                ByVal index As Long)
Const ProcName As String = "SelectChart"
Dim failpoint As Long
On Error GoTo Err

If index > Count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Index must not be less than 1 or greater than Count"
End If

ChartSelector.Tabs(index).Selected = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addFromConfig( _
                ByVal chartSect As ConfigurationSection)
Dim lChart As TradeBuildChart
Dim lTab As MSComctlLib.Tab

Const ProcName As String = "addFromConfig"
Dim failpoint As Long
On Error GoTo Err

Set lChart = loadChartControl
Set lTab = addTab(Nothing)

lChart.LoadFromConfig chartSect, True

lTab.caption = lChart.PeriodLength.ToShortString

fireChange MultiChartAdd

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function addTab( _
                ByVal pPeriodLength As TimePeriod) As MSComctlLib.Tab

Const ProcName As String = "addTab"
Dim failpoint As Long
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function checkIndex( _
                ByVal index As Long) As Long
Const ProcName As String = "checkIndex"
Dim failpoint As Long
On Error GoTo Err

If index = -1 Then
    If mCurrentIndex < 1 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "No current chart"
    Else
        index = mCurrentIndex
    End If
End If

If index > Count Or index < 1 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
            "Index must not be less than 1 or greater than Count"
End If

checkIndex = index

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub closeChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
Const ProcName As String = "closeChart"
Dim failpoint As Long
On Error GoTo Err

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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub fireChange( _
                ByVal changeType As MultiChartChangeTypes)
Dim listener As ChangeListener
Dim ev As ChangeEventData
Const ProcName As String = "fireChange"
Dim failpoint As Long
On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function getChartControlIndexFromIndex(index) As Long
Const ProcName As String = "getChartControlIndexFromIndex"
Dim failpoint As Long
On Error GoTo Err

getChartControlIndexFromIndex = ChartSelector.Tabs(index).Tag

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getChartFromIndex(index) As TradeBuildChart
Const ProcName As String = "getChartFromIndex"
Dim failpoint As Long
On Error GoTo Err

Set getChartFromIndex = TBChart(getChartControlIndexFromIndex(index)).object

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function getIndexFromChartControlIndex(index) As Long
Dim i As Long
Const ProcName As String = "getIndexFromChartControlIndex"
Dim failpoint As Long
On Error GoTo Err

For i = 1 To ChartSelector.Tabs.Count
    If getChartControlIndexFromIndex(i) = index Then
        getIndexFromChartControlIndex = i
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub hideChart( _
                ByVal index As Long)
Dim lChart As TradeBuildChart
Const ProcName As String = "hideChart"
Dim failpoint As Long
On Error GoTo Err

If index = 0 Or index > Count Then Exit Sub
Set lChart = getChartFromIndex(index)
If lChart.State = ChartStateLoaded Then lChart.DisableDrawing
TBChart(getChartControlIndexFromIndex(index)).Visible = False

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub hideTimeframeSelector()
Const ProcName As String = "hideTimeframeSelector"
Dim failpoint As Long
On Error GoTo Err

ControlToolbar.Buttons("selecttimeframe").Width = 0
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = False
resize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function loadChartControl() As TradeBuildChart
Const ProcName As String = "loadChartControl"
On Error GoTo Err

load TBChart(TBChart.UBound + 1)
TBChart(TBChart.UBound).align = vbAlignTop
TBChart(TBChart.UBound).Top = 0
TBChart(TBChart.UBound).Height = ChartSelector.Top
mCount = mCount + 1
Set loadChartControl = TBChart(TBChart.UBound).object

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Function nextIndex( _
                ByVal index As Long) As Long
Const ProcName As String = "nextIndex"
Dim failpoint As Long
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub resize()
Const ProcName As String = "resize"
Dim failpoint As Long
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
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function ShowChart( _
                ByVal index As Long) As TradeBuildChart
Dim lChart As TradeBuildChart
Const ProcName As String = "ShowChart"
Dim failpoint As Long
On Error GoTo Err

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

TimeframeSelector1.selectTimeframe lChart.PeriodLength

Set ShowChart = lChart

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub showTimeframeSelector()
Const ProcName As String = "showTimeframeSelector"
Dim failpoint As Long
On Error GoTo Err

ControlToolbar.Buttons("selecttimeframe").Width = TimeframeSelector1.Width
ControlToolbar.Width = ControlToolbar.Buttons("remove").Left + _
                    ControlToolbar.Buttons("remove").Width
TimeframeSelector1.Visible = True
resize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub storeSettings()
Dim i As Long
Dim lChart As TradeBuildChart
Dim cs As ConfigurationSection

Const ProcName As String = "storeSettings"
Dim failpoint As Long
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

If Not mTicker Is Nothing Then
    mConfig.SetSetting ConfigSettingWorkspace, mTicker.Workspace.name
    mConfig.SetSetting ConfigSettingTickerKey, mTicker.Key
End If

If Not mSpec Is Nothing Then mSpec.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionChartSpecifier)

If Not mStyle Is Nothing Then mConfig.SetSetting ConfigSettingChartStyle, mStyle.name

mConfig.SetSetting ConfigSettingBarFormatterFactoryName, mBarFormatterFactoryName
mConfig.SetSetting ConfigSettingBarFormatterLibraryName, mBarFormatterLibraryName

Set cs = mConfig.AddConfigurationSection(ConfigSectionTradeBuildCharts)
For i = 1 To TBChart.UBound
    If Not TBChart(i) Is Nothing Then
        Set lChart = TBChart(i).object
        lChart.ConfigurationSection = cs.AddConfigurationSection(ConfigSectionTradeBuildChart & "(" & GenerateGUIDString & ")")
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub switchToChart( _
                ByVal index As Long)
Const ProcName As String = "switchToChart"
Dim failpoint As Long
On Error GoTo Err

If index = mCurrentIndex Then Exit Sub

hideChart mCurrentIndex
ShowChart index

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Sub

Private Sub unloadChartControl(ByVal index As Long)
Const ProcName As String = "unloadChartControl"
On Error GoTo Err

Unload TBChart(getChartControlIndexFromIndex(index))
mCount = mCount - 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
