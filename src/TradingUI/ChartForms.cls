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
                ByVal pStyle As ChartStyle)
Const ProcName As String = "Add"
On Error GoTo Err

Dim lChartForm As New fChart
lChartForm.Initialise pDataSource, pPeriodLength, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig, pSpec, pStyle
mChartForms.Add lChartForm

lChartForm.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddFromConfig( _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pTimeframes As Timeframes, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pTimePeriodValidator As ITimePeriodValidator, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "AddFromConfig"
On Error GoTo Err

Dim lChartForm As New fChart
AddFromConfig = lChartForm.LoadFromConfig(pDataSource, pTimeframes, pBarFormatterLibManager, pTimePeriodValidator, pConfig)
If AddFromConfig Then
    mChartForms.Add lChartForm
    lChartForm.Show vbModeless
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
                ByVal pStyle As ChartStyle)
Const ProcName As String = "AddHistoric"
On Error GoTo Err

Dim lChartForm As New fChart
lChartForm.InitialiseHistoric pPeriodLength, pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, pConfig, pSpec, pStyle
mChartForms.Add lChartForm

lChartForm.Show vbModeless

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddHistoricFromConfig( _
                ByVal pContractFuture As IFuture, _
                ByVal pStudyManager As StudyManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pBarFormatterLibManager As BarFormatterLibManager, _
                ByVal pConfig As ConfigurationSection) As Boolean
Const ProcName As String = "AddHistoricFromConfig"
On Error GoTo Err

Dim lChartForm As New fChart
AddHistoricFromConfig = lChartForm.LoadHistoricFromConfig(pContractFuture, pStudyManager, pHistDataStore, pBarFormatterLibManager, pConfig)
If AddHistoricFromConfig Then
    mChartForms.Add lChartForm
    lChartForm.Show vbModeless
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

Dim lChartForm As fChart
For Each lChartForm In mChartForms
    Unload lChartForm
Next
mChartForms.Clear

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




