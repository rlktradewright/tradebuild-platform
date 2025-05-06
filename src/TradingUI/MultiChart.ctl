VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.UserControl MultiChart 
   Alignable       =   -1  'True
   BackColor       =   &H0000FF00&
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9480
   ScaleHeight     =   7710
   ScaleWidth      =   9480
   Begin MSComctlLib.Toolbar ControlToolbar 
      Height          =   330
      Left            =   5880
      TabIndex        =   0
      Top             =   6120
      Width           =   2760
      _ExtentX        =   4868
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
            Key             =   "change"
            Object.ToolTipText     =   "Change the timeframe for the current chart"
            ImageIndex      =   1
            Style           =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Select a new timeframe and add another chart"
            ImageIndex      =   2
            Style           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Object.ToolTipText     =   "Remove current chart"
            ImageIndex      =   3
         EndProperty
      EndProperty
      Begin TradingUI27.TimeframeSelector TimeframeSelector1 
         Height          =   270
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   476
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
            Picture         =   "MultiChart.ctx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MultiChart.ctx":09EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox ChartSelectorPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      ForeColor       =   &H0000FFFF&
      Height          =   135
      Left            =   240
      ScaleHeight     =   135
      ScaleWidth      =   135
      TabIndex        =   4
      Top             =   7080
      Visible         =   0   'False
      Width           =   135
   End
   Begin MSComctlLib.ImageList ChartSelectorImageList 
      Left            =   720
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar ChartSelectorToolbar 
      Align           =   2  'Align Bottom
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   7170
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   953
      ButtonWidth     =   609
      ButtonHeight    =   953
      Style           =   1
      _Version        =   393216
   End
   Begin TradingUI27.MarketChart TBChart 
      Align           =   1  'Align Top
      Height          =   7170
      Index           =   0
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   12647
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

Implements IDeferredAction
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Change(ev As ChangeEventData)
Event ChartStateChanged(ByVal Index As Long, ev As StateChangeEventData)

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

Private mTheme                              As ITheme

Private mIsRaw                              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Set mChangeListeners = New Listeners
ChartSelectorToolbar.Buttons.Clear
TBChart(0).HorizontalScrollBarVisible = False
TBChart(0).Visible = True
gAdjustToolbarPictureSize "WWWW", ChartSelectorPicture
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
hideTimeframeSelector
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Terminate()
Const ProcName As String = "UserControl_Terminate"
gLogger.Log "MultiChart terminated", ProcName, ModuleName, LogLevelDetail
Debug.Print "MultiChart terminated"
End Sub

'@================================================================================
' IDeferredAction Interface Members
'@================================================================================

''
' Brief description
'
' @remarks
'
' @return
'
' @param name
'
' @see
'
'@/
'
Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

resize

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

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub ChartSelectorToolbar_ButtonClick(ByVal Button As Button)
Const ProcName As String = "ChartSelectorToolbar_ButtonClick"
On Error GoTo Err

mCurrentIndex = switchToChart(Button.Index)
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
    If ControlToolbar.Buttons("add").Value = tbrPressed Then
        ControlToolbar.Buttons("change").Value = tbrUnpressed
        showTimeframeSelector
    Else
        hideTimeframeSelector
    End If
Case "CHANGE"
    If ControlToolbar.Buttons("change").Value = tbrPressed Then
        ControlToolbar.Buttons("add").Value = tbrUnpressed
        showTimeframeSelector
    Else
        hideTimeframeSelector
    End If
Case "REMOVE"
    ControlToolbar.Buttons("add").Value = tbrUnpressed
    ControlToolbar.Buttons("change").Value = tbrUnpressed
    Remove mCurrentIndex
End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TBChart_StateChange(Index As Integer, ev As TWUtilities40.StateChangeEventData)
Const ProcName As String = "TBChart_StateChange"
On Error GoTo Err

Index = getButtonIndexFromChartControlIndex(Index)

If Index = mCurrentIndex And ev.State = ChartStates.ChartStateRunning Then
    ControlToolbar.Buttons("change").Enabled = True
End If

RaiseEvent ChartStateChanged(Index, ev)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TBChart_TimePeriodChange(Index As Integer)
Const ProcName As String = "TBChart_TimePeriodChange"
On Error GoTo Err

Dim lButtonInfo As TWChartButtonInfo
With ChartSelectorToolbar.Buttons(getButtonIndexFromChartControlIndex(Index))
    lButtonInfo = .Tag
    lButtonInfo.Caption = TBChart(Index).TimePeriod.ToShortString
    lButtonInfo.ToolTipText = generateTooltipText(TBChart(Index).TimePeriod)
    .Tag = lButtonInfo
    .ToolTipText = lButtonInfo.ToolTipText
End With

updateButtonImages

fireChange MultiChartPeriodLengthChanged

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub TimeframeSelector1_Click()
Const ProcName As String = "TimeframeSelector1_Click"
On Error GoTo Err

If ControlToolbar.Buttons("add").Value = tbrPressed Then
    Add TimeframeSelector1.TimePeriod
    hideTimeframeSelector
    ControlToolbar.Buttons("add").Value = tbrUnpressed
Else
    Chart.ChangeTimePeriod TimeframeSelector1.TimePeriod
    hideTimeframeSelector
    ControlToolbar.Buttons("change").Value = tbrUnpressed
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
                Optional ByVal Index As Long = -1) As ChartController
Const ProcName As String = "BaseChartController"
On Error GoTo Err

Index = checkIndex(Index)
Set BaseChartController = getChartFromButtonIndex(Index).BaseChartController

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' do not make this Public because the Value returned cannot be handled by non-friend
' components
Friend Property Get Chart( _
                Optional ByVal Index As Long = -1) As MarketChart
Const ProcName As String = "Chart"
On Error GoTo Err

Index = checkIndex(Index)
Set Chart = getChartFromButtonIndex(Index)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ChartManager( _
                Optional ByVal Index As Long = -1) As ChartManager
Const ProcName As String = "ChartManager"
On Error GoTo Err

Index = checkIndex(Index)
Set ChartManager = getChartFromButtonIndex(Index).ChartManager

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

Public Property Get CurrentIndex() As Long
CurrentIndex = mCurrentIndex
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If mConfig Is Value Then Exit Property
If Not mConfig Is Nothing Then mConfig.Remove
If Value Is Nothing Then Exit Property

Set mConfig = Value

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
                ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled = Value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LoadingText( _
                Optional ByVal Index As Long = -1) As Text
Const ProcName As String = "LoadingText"
On Error GoTo Err

Index = checkIndex(Index)
Set LoadingText = getChartFromButtonIndex(Index).LoadingText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Get PriceRegion( _
                Optional ByVal Index As Long = -1) As ChartRegion
Const ProcName As String = "PriceRegion"
On Error GoTo Err

Index = checkIndex(Index)
Set PriceRegion = getChartFromButtonIndex(Index).PriceRegion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Style(ByVal pStyle As ChartStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = pStyle

If Not mConfig Is Nothing Then
    If mStyle Is Nothing Then
        mConfig.SetSetting ConfigSettingChartStyle, ""
    Else
        mConfig.SetSetting ConfigSettingChartStyle, mStyle.Name
    End If
End If

Dim i As Long
For i = 1 To TBChart.UBound
    Dim lChart As MarketChart
    Set lChart = TBChart(i).object
    lChart.BaseChartController.Style = mStyle
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get State( _
                Optional ByVal Index As Long = -1) As ChartStates
Const ProcName As String = "State"
On Error GoTo Err

Index = checkIndex(Index)
State = getChartFromButtonIndex(Index).State

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.ToolbarBackColor
TBChart(0).ChartBackColor = mTheme.ToolbarBackColor
gApplyTheme mTheme, UserControl.Controls

ChartSelectorPicture.BackColor = mTheme.ToolbarBackColor
ChartSelectorPicture.ForeColor = mTheme.TabstripForeColor

updateButtonImages
        
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Get Timeframe( _
                Optional ByVal Index As Long = -1) As Timeframe
Const ProcName As String = "Timeframe"
On Error GoTo Err

Index = checkIndex(Index)
Set Timeframe = getChartFromButtonIndex(Index).Timeframe

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TimePeriod( _
                Optional ByVal Index As Long = -1) As TimePeriod
Const ProcName As String = "TimePeriod"
On Error GoTo Err

Index = checkIndex(Index)
Set TimePeriod = getChartFromButtonIndex(Index).TimePeriod

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get VolumeRegion( _
                Optional ByVal Index As Long = -1) As ChartRegion
Const ProcName As String = "VolumeRegion"
On Error GoTo Err

Index = checkIndex(Index)
Set VolumeRegion = getChartFromButtonIndex(Index).VolumeRegion

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
                Optional ByVal pUpdatePerTick As Boolean = True, _
                Optional ByVal pInitialNumberOfBars As Long = -1, _
                Optional ByVal pIncludeBarsOutsideSession As Boolean = False, _
                Optional ByVal pDeferStart As Boolean) As Long
Const ProcName As String = "Add"
On Error GoTo Err

Assert Not mTimeframes Is Nothing, "Can't add non-raw charts to this MultiChart"

Dim lChartSpec As ChartSpecifier
Set lChartSpec = CreateChartSpecifier( _
                        IIf(pInitialNumberOfBars = -1, mSpec.InitialNumberOfBars, pInitialNumberOfBars), _
                        IIf(pIncludeBarsOutsideSession, True, mSpec.IncludeBarsOutsideSession), _
                        mSpec.FromTime, _
                        mSpec.ToTime, _
                        mSpec.CustomSessionStartTime, _
                        mSpec.CustomSessionEndTime)

Dim lChart As MarketChart
Set lChart = loadChartControl

Dim lButton As MSComctlLib.Button
Set lButton = addChartSelectorButton(pPeriodLength)
gLogger.Log "Added button with index " & lButton.Index, ProcName, ModuleName

' we notify the add before calling lChart.ShowChart so that it's before
' the ChartStates.ChartStateRunning event
fireChange MultiChartAdd

lChart.SetupChart pDeferStart, mTimeframes, pPeriodLength, lChartSpec, mStyle, pUpdatePerTick, mBarFormatterLibManager, mBarFormatterFactoryName, mBarFormatterLibraryName, mExcludeCurrentBar, pTitle
resize

If Not mConfig Is Nothing Then
    lChart.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionMarketCharts).AddConfigurationSection(ConfigSectionMarketChart & "(" & GenerateGUIDString & ")")
End If

If Not pDeferStart Then
    If mCurrentIndex > 0 Then ChartSelectorToolbar.Buttons(mCurrentIndex).Value = tbrUnpressed
    lButton.Value = tbrPressed
    
    mCurrentIndex = switchToChart(lButton.Index)
    Add = mCurrentIndex
Else
    Add = lButton.Index
End If

fireChange MultiChartSelectionChanged

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

' return the index rather than the chart because the chart cannot be handled by non-friend
' components
Public Function AddRaw( _
                ByVal pTimeframe As Timeframe, _
                ByVal pStudyManager As StudyManager, _
                Optional ByVal pLocalSymbol As String, _
                Optional ByVal pSecType As SecurityTypes, _
                Optional ByVal pExchange As String, _
                Optional ByVal pTickSize As Double, _
                Optional ByVal pSessionStartTime As Date, _
                Optional ByVal pSessionEndTime As Date, _
                Optional ByVal pTitle As String, _
                Optional ByVal pUpdatePerTick As Boolean = True) As Long
Const ProcName As String = "AddRaw"
On Error GoTo Err

Assert mTimeframes Is Nothing, "Can't add raw charts to this MultiChart"

Dim lChart As MarketChart
Set lChart = loadChartControl

Dim lButton As MSComctlLib.Button
Set lButton = addChartSelectorButton(pTimeframe.TimePeriod)

' we notify the add before calling lChart.ShowChartRaw so that it's before
' the ChartStates.ChartStateRunning event
fireChange MultiChartAdd

lChart.ShowChartRaw pStudyManager, _
                    pTimeframe, _
                    mStyle, _
                    pUpdatePerTick, _
                    pLocalSymbol, _
                    pSecType, _
                    pExchange, _
                    pTickSize, _
                    pSessionStartTime, _
                    pSessionEndTime, _
                    mBarFormatterLibManager, _
                    mBarFormatterFactoryName, _
                    mBarFormatterLibraryName, _
                    pTitle

If mCurrentIndex > 0 Then ChartSelectorToolbar.Buttons(mCurrentIndex).Value = tbrUnpressed
lButton.Value = tbrPressed
mCurrentIndex = switchToChart(lButton.Index)
AddRaw = mCurrentIndex

fireChange MultiChartSelectionChanged

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddChangeListener( _
                ByVal pListener As IChangeListener)
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

Do While ChartSelectorToolbar.Buttons.Count <> 0
    Remove ChartSelectorToolbar.Buttons.Count
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FocusChart()
Const ProcName As String = "FocusChart"
On Error GoTo Err

Dim Index As Long
Index = checkIndex(-1)
TBChart(getChartControlIndexFromButtonIndex(Index)).SetFocus

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetIndexFromTimeframe(ByVal pTimeframe As Timeframe) As Long
Const ProcName As String = "GetIndexFromTimeframe"
On Error GoTo Err

Dim i As Long
For i = 1 To Count
    If Timeframe(i).TimePeriod Is pTimeframe.TimePeriod Then
        GetIndexFromTimeframe = i
        Exit Function
    End If
Next

AssertArgument False, "pTimeframe not found"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim i As Long
For i = 1 To mCount
    getChartFromButtonIndex(i).Finish
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

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManager must be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "pBarFormatterLibraryName and pBarFormatterFactoryName must both be blank or non-blank"

Set mTimeframes = pTimeframes

Set mSpec = pSpec

If pStyle Is Nothing Then
    Set mStyle = ChartStylesManager.DefaultStyle
Else
    Set mStyle = pStyle
End If

TBChart(0).ChartBackColor = pBackColor
mIsHistoric = (mSpec.ToTime <> 0)

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

Public Sub InitialiseRaw( _
                Optional ByVal pStyle As ChartStyle, _
                Optional ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pBarFormatterFactoryName As String, _
                Optional ByVal pBarFormatterLibraryName As String, _
                Optional ByVal pBackColor As Long = &HC0C0C0)
Const ProcName As String = "InitialiseRaw"
On Error GoTo Err

AssertArgument pBarFormatterFactoryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterFactoryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument pBarFormatterLibraryName = "" Or Not pBarFormatterLibManager Is Nothing, "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"
AssertArgument (pBarFormatterLibraryName = "" And pBarFormatterFactoryName = "") Or (pBarFormatterLibraryName <> "" And pBarFormatterFactoryName <> ""), "If pBarFormatterLibraryName is not blank then pBarFormatterLibManagermust be supplied"

mIsRaw = True

If pStyle Is Nothing Then
    Set mStyle = ChartStylesManager.DefaultStyle
Else
    Set mStyle = pStyle
End If

TBChart(0).ChartBackColor = pBackColor

Set mBarFormatterLibManager = pBarFormatterLibManager
mBarFormatterFactoryName = pBarFormatterFactoryName
mBarFormatterLibraryName = pBarFormatterLibraryName

ControlToolbar.Visible = False

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

mIsHistoric = (mSpec.ToTime <> 0)

Dim lCurrentChartIndex As Long
lCurrentChartIndex = CLng(mConfig.GetSetting(ConfigSettingCurrentChart, "1"))

Dim cs As ConfigurationSection
For Each cs In mConfig.AddConfigurationSection(ConfigSectionMarketCharts)
    AddFromConfig cs
Next

If ChartSelectorToolbar.Buttons.Count > 0 Then SelectChart lCurrentChartIndex

resize

LoadFromConfig = True

Exit Function

Err:
gHandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pModuleName:=ModuleName
LoadFromConfig = False
End Function

Public Sub Remove( _
                ByVal Index As Long)
Const ProcName As String = "Remove"
On Error GoTo Err

AssertArgument Index <= Count And Index >= 1, "Index must not be less than 1 or greater than Count"

Dim nxtIndex As Long
nxtIndex = nextIndex(Index)
closeChart Index
If Index = mCurrentIndex Then mCurrentIndex = 0

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

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)
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
                ByVal Index As Long)
Const ProcName As String = "SelectChart"
On Error GoTo Err

AssertArgument Index <= Count And Index >= 1, "Index must not be less than 1 or greater than Count"

ChartSelectorToolbar.Buttons.Item(Index).Value = tbrPressed
mCurrentIndex = switchToChart(Index)
resize
fireChange MultiChartSelectionChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetStudyManager( _
                ByVal Value As StudyManager, _
                Optional ByVal Index As Long = -1)
Const ProcName As String = "SetStudyManager"
On Error GoTo Err

Index = checkIndex(Index)
getChartFromButtonIndex(Index).StudyManager = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UpdateLastBar()
Const ProcName As String = "UpdateLastBar"
On Error GoTo Err

Dim i As Long
For i = 1 To ChartSelectorToolbar.Buttons.Count
    Dim lChart As MarketChart
    Set lChart = TBChart(getChartControlIndexFromButtonIndex(i)).object
    lChart.UpdateLastBar
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addChartSelectorButton( _
                ByVal pPeriodLength As TimePeriod) As Button
Const ProcName As String = "addChartSelectorButton"
On Error GoTo Err

If ChartSelectorToolbar.Buttons.Count = 0 Then
    ' adding the first button causes the Toolbar to be resized, but this
    ' happens asynchronously after the current event handler exit,
    ' so do a deferred resize to take account of this
    DeferAction Me
End If

If pPeriodLength Is Nothing Then
    Set addChartSelectorButton = ChartSelectorToolbar.Buttons.Add(, , "")
Else
    Dim lButtonInfo As TWChartButtonInfo
    gCreateButtonInfo lButtonInfo, pPeriodLength.ToShortString, tbrButtonGroup, tbrUnpressed, generateTooltipText(pPeriodLength), True, TBChart.UBound
    gAddButtonImageToImageList ChartSelectorImageList, lButtonInfo, ChartSelectorPicture
    If ChartSelectorImageList.ListImages.Count = 1 Then Set ChartSelectorToolbar.ImageList = ChartSelectorImageList
    Set addChartSelectorButton = gAddButtonToToolbar(ChartSelectorToolbar, lButtonInfo)
End If

ControlToolbar.Buttons("remove").Enabled = True

If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCurrentChart, ChartSelectorToolbar.Buttons.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub AddFromConfig( _
                ByVal chartSect As ConfigurationSection)
Const ProcName As String = "AddFromConfig"
On Error GoTo Err

Dim lChart As MarketChart
Set lChart = loadChartControl
lChart.LoadFromConfig mTimeframes, chartSect, mBarFormatterLibManager, True

addChartSelectorButton lChart.TimePeriod

fireChange MultiChartAdd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function checkIndex( _
                ByVal Index As Long) As Long
Const ProcName As String = "checkIndex"
On Error GoTo Err

If Index = -1 Then
    Assert mCurrentIndex >= 1, "No current chart"
    Index = mCurrentIndex
End If

AssertArgument Index <= Count And Index >= 1, "Index must not be less than 1 or greater than Count"

checkIndex = Index

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub closeChart( _
                ByVal Index As Long)
Const ProcName As String = "closeChart"
On Error GoTo Err

Dim lChart As MarketChart
Set lChart = getChartFromButtonIndex(Index)
lChart.RemoveFromConfig
lChart.Finish
unloadChartControl Index
ChartSelectorToolbar.Buttons.Remove Index
If ChartSelectorToolbar.Buttons.Count = 0 Then
    ControlToolbar.Buttons("remove").Enabled = False
    TBChart(0).Visible = True
    TBChart(0).Height = ChartSelectorToolbar.Top
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

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function displayChart( _
                ByVal Index As Long) As MarketChart
Const ProcName As String = "displayChart"
On Error GoTo Err

If Index = 0 Then Exit Function

gLogger.Log "Showing chart " & Index, ProcName, ModuleName

Dim lChart As MarketChart: Set lChart = getChartFromButtonIndex(Index)

If lChart.State = ChartStates.ChartStateCreated Then lChart.Start

If lChart.State = ChartStateRunning Then
    gLogger.Log "EnableDrawing", ProcName, ModuleName, LogLevelDetail
    lChart.EnableDrawing
    ControlToolbar.Buttons("change").Enabled = True
Else
    ControlToolbar.Buttons("change").Enabled = False
End If

TBChart(getChartControlIndexFromButtonIndex(Index)).Align = AlignConstants.vbAlignTop
TBChart(getChartControlIndexFromButtonIndex(Index)).Height = ChartSelectorToolbar.Top
TBChart(getChartControlIndexFromButtonIndex(Index)).Visible = True

If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCurrentChart, Index

If Not mIsRaw Then TimeframeSelector1.SelectTimeframe lChart.TimePeriod

Set displayChart = lChart

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateTooltipText(ByVal pPeriodLength As TimePeriod) As String
generateTooltipText = "Switch to " & pPeriodLength.ToString & " chart"
End Function

Private Function getButtonIndexFromChartControlIndex(Index) As Long
Const ProcName As String = "getButtonIndexFromChartControlIndex"
On Error GoTo Err

Dim i As Long
For i = 1 To ChartSelectorToolbar.Buttons.Count
    If getChartControlIndexFromButtonIndex(i) = Index Then
        getButtonIndexFromChartControlIndex = i
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getChartControlIndexFromButtonIndex(Index) As Long
Const ProcName As String = "getChartControlIndexFromButtonIndex"
On Error GoTo Err

Dim l As TWChartButtonInfo
l = ChartSelectorToolbar.Buttons(Index).Tag
getChartControlIndexFromButtonIndex = l.ChartIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getChartFromButtonIndex(Index) As MarketChart
Const ProcName As String = "getChartFromButtonIndex"
On Error GoTo Err

Set getChartFromButtonIndex = TBChart(getChartControlIndexFromButtonIndex(Index)).object

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideChart( _
                ByVal Index As Long)
Const ProcName As String = "hideChart"
On Error GoTo Err

If Index = 0 Or Index > Count Then Exit Sub

gLogger.Log "Hiding chart " & Index, ProcName, ModuleName

Dim lChart As MarketChart: Set lChart = getChartFromButtonIndex(Index)

gLogger.Log "DisableDrawing", ProcName, ModuleName, LogLevelHighDetail
lChart.DisableDrawing

TBChart(getChartControlIndexFromButtonIndex(Index)).Visible = False
TBChart(getChartControlIndexFromButtonIndex(Index)).Align = vbAlignNone

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
TBChart(TBChart.UBound).Height = ChartSelectorToolbar.Top
mCount = mCount + 1
Set loadChartControl = TBChart(TBChart.UBound).object

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function nextIndex( _
                ByVal Index As Long) As Long
Const ProcName As String = "nextIndex"
On Error GoTo Err

If Index > 1 Then
    nextIndex = Index - 1
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

ControlToolbar.ZOrder 0

If Count > 0 And mCurrentIndex > 0 Then
    TBChart(getChartControlIndexFromButtonIndex(mCurrentIndex)).Height = ChartSelectorToolbar.Top
Else
    TBChart(0).Height = ChartSelectorToolbar.Top
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Private Function switchToChart( _
                ByVal Index As Long) As Long
Const ProcName As String = "switchToChart"
On Error GoTo Err

switchToChart = Index

If Index = mCurrentIndex Then
    gLogger.Log "Chart " & Index & " already selected", ProcName, ModuleName
    Exit Function
End If

gLogger.Log "Selecting chart " & Index, ProcName, ModuleName
gLogger.Log "Hide chart " & mCurrentIndex, ProcName, ModuleName
hideChart mCurrentIndex
displayChart Index

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub unloadChartControl(ByVal Index As Long)
Const ProcName As String = "unloadChartControl"
On Error GoTo Err

Unload TBChart(getChartControlIndexFromButtonIndex(Index))
mCount = mCount - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateButtonImages()
Const ProcName As String = "updateButtonImages"
On Error GoTo Err

Dim failpoint As String
failpoint = "100"

If ChartSelectorToolbar.Buttons.Count = 0 Then Exit Sub

Set ChartSelectorToolbar.ImageList = Nothing

Dim i As Long
For i = 1 To ChartSelectorToolbar.Buttons.Count
    Dim lButtonInfo As TWChartButtonInfo
    failpoint = "200"
    With ChartSelectorToolbar.Buttons(i)
        failpoint = "300"
        lButtonInfo = .Tag
        failpoint = "400"
        gUpdateButtonImageInImageList ChartSelectorImageList, lButtonInfo, ChartSelectorPicture
    End With
Next

failpoint = "500"
Set ChartSelectorToolbar.ImageList = ChartSelectorImageList

For i = 1 To ChartSelectorToolbar.Buttons.Count
    With ChartSelectorToolbar.Buttons(i)
        .Image = Empty
        .Image = .Key
    End With
Next
        
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failpoint
End Sub
