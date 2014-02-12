VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ChartFormController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "ChartFormController"

'@================================================================================
' Member variables
'@================================================================================

Private mChartForm                                  As fChart

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mChartForm = New fChart

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get chartForm() As fChart
Set chartForm = mChartForm
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "Initialise"
On Error GoTo Err

mChartForm.Initialise pContractFuture, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseHistoric( _
                ByVal pContractFuture As IFuture, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "InitialiseHistoric"
On Error GoTo Err

mChartForm.InitialiseHistoric pContractFuture, pHistDataStore, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithDataSource( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "InitialiseWithDataSource"
On Error GoTo Err

mChartForm.InitialiseWithDataSource pDataSource, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseWithTicker( _
                ByVal pTicker As Ticker, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection, _
                ByVal pSpec As ChartSpecifier, _
                ByVal pStyle As ChartStyle)
Const ProcName As String = "InitialiseWithTicker"
On Error GoTo Err

mChartForm.InitialiseWithTicker pTicker, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function LoadFromConfig( _
                ByVal pContractFuture As IFuture, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

LoadFromConfig = mChartForm.LoadFromConfig(pContractFuture, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function LoadFromConfigWithDataSource( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "LoadFromConfigWithDataSource"
On Error GoTo Err

LoadFromConfigWithDataSource = mChartForm.LoadFromConfigWithDataSource(pDataSource, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function LoadFromConfigWithTicker( _
                ByVal pTicker As Ticker, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "LoadFromConfigWithTicker"
On Error GoTo Err

LoadFromConfigWithTicker = mChartForm.LoadFromConfigWithTicker(pTicker, pBarFormatterLibManager, pTimePeriodValidator, pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ShowChart( _
                ByVal pPeriodLength As TimePeriod)
Const ProcName As String = "showChart"
On Error GoTo Err

mChartForm.ShowChart pPeriodLength

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Show(ByVal pModal As FormShowConstants)
Const ProcName As String = "Show"
On Error GoTo Err

mChartForm.Show pModal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UnloadChart()
Const ProcName As String = "UnloadChart"
On Error GoTo Err

Unload mChartForm

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




