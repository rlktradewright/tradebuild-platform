VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartForms"
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

Private Const ModuleName                            As String = "ChartForms"

'@================================================================================
' Member variables
'@================================================================================

Private mChartForms                                 As New EnumerableCollection

Private mTheme                                      As ITheme

Private mChartsVisible                              As Boolean
Private mHistoricalChartsVisible                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator() As Enumerator
Set Enumerator = mChartForms.Enumerator
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

Dim l As fChart
For Each l In mChartForms
    If (l.IsHistorical And mHistoricalChartsVisible) Or _
        (Not l.IsHistorical And mChartsVisible) Then
        LogMessage "Applying theme to form " & l.Caption, LogLevelHighDetail
        l.Theme = mTheme
    End If
Next

gSetStudyPickerTheme mTheme

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

Public Sub Add( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme)
Const ProcName As String = "Add"
On Error GoTo Err

Dim lChartForm As fChart
Set lChartForm = CreateChart(pDataSource, pPeriodLength, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle, pOwner)

If Not IsMissing(pOwner) Then lChartForm.Show vbModeless, pOwner
If Not pTheme Is Nothing Then lChartForm.Theme = pTheme

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddAsync( _
                ByVal pTickers As SelectedTickers, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme, _
                Optional ByVal pPriority As TaskPriorities, _
                Optional ByVal pName As String, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "AddAsync"
On Error GoTo Err

Dim t As New ChartFormsAddTask
t.Initialise Me, pTickers, pPeriodLength, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle, pOwner, pTheme
Set AddAsync = StartTask(t, pPriority, pName, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddFromConfig( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme) As Boolean
Const ProcName As String = "AddFromConfig"
On Error GoTo Err

Dim lChartForm As New fChart
AddFromConfig = lChartForm.LoadFromConfig(pDataSource, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pOwner)
If AddFromConfig Then
    mChartForms.Add lChartForm
    If Not IsMissing(pOwner) Then lChartForm.Show vbModeless, pOwner
    If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
Else
    Unload lChartForm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddHistoric( _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme)
Const ProcName As String = "AddHistoric"
On Error GoTo Err

Dim lChartForm As New fChart
Set lChartForm = CreateHistoricalChart(pPeriodLength, pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, pConfig, pSpec, pStyle, pOwner)

If Not IsMissing(pOwner) Then lChartForm.Show vbModeless, pOwner
If Not pTheme Is Nothing Then lChartForm.Theme = pTheme

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddHistoricAsync( _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pContracts As IContracts, _
                ByVal pStudyLibManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme, _
                Optional ByVal pPriority As TaskPriorities, _
                Optional ByVal pName As String, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "AddHistoricAsync"
On Error GoTo Err

Dim t As New ChartFormsAddTask
t.InitialiseHistorical Me, pPeriodLength, pContracts, pStudyLibManager, pHistDataStore, pBarFormatterLibManager, pConfig, pSpec, pStyle, pOwner, pTheme
Set AddHistoricAsync = StartTask(t, pPriority, pName, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddHistoricFromConfig( _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme) As Boolean
Const ProcName As String = "AddHistoricFromConfig"
On Error GoTo Err

Dim lChartForm As New fChart
AddHistoricFromConfig = lChartForm.LoadHistoricFromConfig(pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, pConfig)
If AddHistoricFromConfig Then
    mChartForms.Add lChartForm
    If Not IsMissing(pOwner) Then lChartForm.Show vbModeless, pOwner
    If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
Else
    Unload lChartForm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateChart( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                ByVal pOwner As Variant) As fChart
Const ProcName As String = "CreateChart"
On Error GoTo Err

Dim lConfig As ConfigurationSection

If Not pConfig Is Nothing Then
    Set lConfig = pConfig.AddConfigurationSection(ConfigSectionChart & "(" & GenerateGUIDString & ")")
    lConfig.SetSetting ConfigSettingDataSourceKey, pDataSource.Key
End If

Dim lChartForm As New fChart
lChartForm.Initialise pDataSource, pPeriodLength, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, lConfig, pSpec, pStyle, pOwner
mChartForms.Add lChartForm

Set CreateChart = lChartForm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateChartFromConfig( _
                ByVal pChartConfig As ConfigurationSection, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator) As Boolean
Const ProcName As String = "CreateChartFromConfig"
On Error GoTo Err

Dim lTickerKey As String
lTickerKey = pChartConfig.GetSetting(ConfigSettingDataSourceKey, "")

Dim lTicker As Ticker
If lTickerKey = "" Then
    Set lTicker = pTickers.GetTicker(pChartConfig.InstanceQualifier)
Else
    Set lTicker = pTickers.GetTicker(lTickerKey)
End If

If lTicker Is Nothing Then
    CreateChartFromConfig = False
ElseIf Not (lTicker.State = MarketDataSourceStateReady Or _
        lTicker.State = MarketDataSourceStatePaused Or _
        lTicker.State = MarketDataSourceStateRunning) Then
    CreateChartFromConfig = False
Else
    CreateChartFromConfig = AddFromConfig(lTicker, _
                    lTicker.Timeframes, _
                    pBarFormatterLibManager, _
                    pTimePeriodValidator, _
                    pChartConfig, _
                    , _
                    Nothing)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateHistoricalChart( _
                ByVal pPeriodLength As TimePeriod, _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle, _
                Optional ByVal pOwner As Variant) As fChart
Const ProcName As String = "CreateHistoricalChart"
On Error GoTo Err

Dim lConfig As ConfigurationSection
If Not pConfig Is Nothing Then
    Set lConfig = pConfig.AddConfigurationSection(ConfigSectionChart & "(" & GenerateGUIDString & ")")
End If
    
Dim lChartForm As New fChart
lChartForm.InitialiseHistoric pPeriodLength, pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, lConfig, pSpec, pStyle, pOwner
mChartForms.Add lChartForm

Set CreateHistoricalChart = lChartForm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function CreateHistoricalChartFromConfig( _
                ByVal pChartConfig As ConfigurationSection, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager) As Boolean
Const ProcName As String = "CreateHistoricalChartFromConfig"
On Error GoTo Err

Dim lContract As IContract
Set lContract = LoadContractFromConfig(pChartConfig.GetConfigurationSection(ConfigSectionContract))

CreateHistoricalChartFromConfig = AddHistoricFromConfig(CreateFuture(lContract), _
                                                    pStudyLibraryManager.CreateStudyManager, _
                                                    pHistDataStore, _
                                                    pBarFormatterLibManager, _
                                                    pChartConfig, _
                                                    , _
                                                    Nothing)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim f As fChart
For Each f In mChartForms
    f.Hide
Next

For Each f In mChartForms
    gLogger.Log "Closing form: caption=" & f.Caption & "; type=" & TypeName(f), ProcName, ModuleName
    Unload f
Next
mChartForms.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub HideCharts()
Const ProcName As String = "HideCharts"
On Error GoTo Err

mChartsVisible = False

Dim lChart As fChart
For Each lChart In mChartForms
    If Not lChart.IsHistorical Then lChart.Hide
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub HideHistoricalCharts()
Const ProcName As String = "HideHistoricalCharts"
On Error GoTo Err

mHistoricalChartsVisible = False

Dim lChart As fChart
For Each lChart In mChartForms
    If lChart.IsHistorical Then lChart.Hide
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadChartsFromConfig( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme)
Const ProcName As String = "LoadChartsFromConfig"
On Error GoTo Err

Dim lRemovableCSs As New Collection

Dim chartConfig As ConfigurationSection
For Each chartConfig In pConfig
    If Not CreateChartFromConfig(chartConfig, pTickers, pBarFormatterLibManager, pTimePeriodValidator) Then lRemovableCSs.Add chartConfig
Next

For Each chartConfig In lRemovableCSs
    chartConfig.Remove
Next

If Not IsMissing(pOwner) Then
    Dim lChart As fChart
    For Each lChart In mChartForms
        If Not lChart.IsHistorical Then
            lChart.Show vbModeless, pOwner
            If Not pTheme Is Nothing Then lChart.Theme = pTheme
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function LoadChartsFromConfigAsync( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme, _
                Optional ByVal pPriority As TaskPriorities, _
                Optional ByVal pName As String, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "LoadChartsFromConfigAsync"
On Error GoTo Err

Dim t As New ChartFormsLoadTask
t.Initialise Me, pConfig, pTickers, pBarFormatterLibManager, pTimePeriodValidator, pOwner, pTheme
Set LoadChartsFromConfigAsync = StartTask(t, pPriority, pName, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Public Sub LoadHistoricalChartsFromConfig( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme)
Const ProcName As String = "LoadHistoricalChartsFromConfig"
On Error GoTo Err

Dim chartConfig As ConfigurationSection
For Each chartConfig In pConfig
    CreateHistoricalChartFromConfig chartConfig, pStudyLibraryManager, pHistDataStore, pBarFormatterLibManager
Next

If Not IsMissing(pOwner) Then
    Dim lChart As fChart
    For Each lChart In mChartForms
        If lChart.IsHistorical Then
            lChart.Show vbModeless, pOwner
            If Not pTheme Is Nothing Then lChart.Theme = pTheme
        End If
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function LoadHistoricalChartsFromConfigAsync( _
                ByVal pConfig As ConfigurationSection, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                Optional ByVal pOwner As Variant, _
                Optional ByVal pTheme As ITheme, _
                Optional ByVal pPriority As TaskPriorities, _
                Optional ByVal pName As String, _
                Optional ByVal pCookie As Variant) As TaskController
Const ProcName As String = "LoadHistoricalChartsFromConfigAsync"
On Error GoTo Err

Dim t As New ChartFormsLoadTask
t.InitialiseHistorical Me, pConfig, pStudyLibraryManager, pHistDataStore, pBarFormatterLibManager, pOwner, pTheme
Set LoadHistoricalChartsFromConfigAsync = StartTask(t, pPriority, pName, pCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


Public Sub SetStyle(ByVal pStyle As ChartStyle, ByVal pHistorical As Boolean)
Const ProcName As String = "SetStyle"
On Error GoTo Err

Dim lChartForm As fChart
For Each lChartForm In mChartForms
    If (lChartForm.IsHistorical And pHistorical) Or _
       Not (lChartForm.IsHistorical Or pHistorical) Then lChartForm.Style = pStyle
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowCharts(ByVal pOwner As Object)
Const ProcName As String = "ShowCharts"
On Error GoTo Err

mChartsVisible = True

Dim lChart As fChart
For Each lChart In mChartForms
    If Not lChart.IsHistorical Then
        lChart.Show vbModeless, pOwner
        If Not mTheme Is Nothing Then lChart.Theme = mTheme
        lChart.Owner = pOwner
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ShowHistoricalCharts(ByVal pOwner As Object)
Const ProcName As String = "ShowHistoricalCharts"
On Error GoTo Err

mHistoricalChartsVisible = True

Dim lChart As fChart
For Each lChart In mChartForms
    If lChart.IsHistorical Then
        lChart.Show vbModeless, pOwner
        If Not mTheme Is Nothing Then lChart.Theme = mTheme
        lChart.Owner = pOwner
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




