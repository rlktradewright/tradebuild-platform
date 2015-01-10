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

Public Property Let Theme(ByVal value As ITheme)
Set mTheme = value

Dim l As fChart
For Each l In mChartForms
    l.Theme = mTheme
Next

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

Dim lConfig As ConfigurationSection

If Not pConfig Is Nothing Then
    Set lConfig = pConfig.AddConfigurationSection(ConfigSectionChart & "(" & GenerateGUIDString & ")")
    lConfig.SetSetting ConfigSettingDataSourceKey, pDataSource.Key
End If

Dim lChartForm As New fChart
lChartForm.Initialise pDataSource, pPeriodLength, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, lConfig, pSpec, pStyle, pOwner
mChartForms.Add lChartForm

If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
lChartForm.Show vbModeless, pOwner

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
    lChartForm.Show vbModeless, pOwner
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

Dim lConfig As ConfigurationSection
If Not pConfig Is Nothing Then
    Set lConfig = pConfig.AddConfigurationSection(ConfigSectionChart & "(" & GenerateGUIDString & ")")
End If

    
Dim lChartForm As New fChart
lChartForm.InitialiseHistoric pPeriodLength, pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, lConfig, pSpec, pStyle, pOwner
mChartForms.Add lChartForm

If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
lChartForm.Show vbModeless, pOwner

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    If Not pTheme Is Nothing Then lChartForm.Theme = pTheme
    lChartForm.Show vbModeless, pOwner
Else
    Unload lChartForm
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim f As fChart
For Each f In mChartForms
    gLogger.Log "Closing form: caption=" & f.caption & "; type=" & TypeName(f), ProcName, ModuleName
    Unload f
Next
mChartForms.Clear

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
    If Not createChartFromConfig(chartConfig, pTickers, pBarFormatterLibManager, pTimePeriodValidator, pOwner, pTheme) Then lRemovableCSs.Add chartConfig
Next

For Each chartConfig In lRemovableCSs
    chartConfig.Remove
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
    createHistoricalChartFromConfig chartConfig, pStudyLibraryManager, pHistDataStore, pBarFormatterLibManager, pOwner, pTheme
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

'@================================================================================
' Helper Functions
'@================================================================================

Private Function createChartFromConfig( _
                ByVal pChartConfig As ConfigurationSection, _
                ByVal pTickers As Tickers, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme) As Boolean
Const ProcName As String = "createChartFromConfig"
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
    createChartFromConfig = False
Else
    createChartFromConfig = AddFromConfig(lTicker, _
                    lTicker.Timeframes, _
                    pBarFormatterLibManager, _
                    pTimePeriodValidator, _
                    pChartConfig, _
                    pOwner, _
                    pTheme)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createHistoricalChartFromConfig( _
                ByVal pChartConfig As ConfigurationSection, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pOwner As Variant, _
                ByVal pTheme As ITheme) As Boolean
Const ProcName As String = "createHistoricalChartFromConfig"
On Error GoTo Err

Dim lContract As IContract
Set lContract = LoadContractFromConfig(pChartConfig.GetConfigurationSection(ConfigSectionContract))

createHistoricalChartFromConfig = AddHistoricFromConfig(CreateFuture(lContract), _
                                                    pStudyLibraryManager.CreateStudyManager, _
                                                    pHistDataStore, _
                                                    pBarFormatterLibManager, _
                                                    pChartConfig, _
                                                    pOwner, _
                                                    pTheme)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




