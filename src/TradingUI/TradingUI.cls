VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingUI"
Attribute VB_GlobalNameSpace = True
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

Public Enum ChartStates
    ''
    ' No chart has yet been started, or the chart has been cleared.
    '
    '@/
    ChartStateBlank

    ''
    ' The chart has been created, historical data (if any) has not
    ' yet been requested.
    '
    '@/
    ChartStateCreated

    ''
    ' Historic data is being fetched.
    '
    '@/
    ChartStateFetching

    ''
    ' Historic data has been fetched and is being added to the chart.
    '
    '@/
    ChartStateLoading
    
    ''
    ' All historic data (if any) has been added to the chart and the
    ' chart is now running normally.
    '
    '@/
    ChartStateRunning
End Enum

Public Enum MultiChartChangeTypes
    MultiChartSelectionChanged
    MultiChartAdd
    MultiChartRemove
    MultiChartPeriodLengthChanged
End Enum

Public Enum OrderTicketModes
    OrderTicketModeNone
    OrderTicketModeLiveOnly
    OrderTicketModeSimulatedOnly
    OrderTicketModeLiveAndSimulated
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TWChartButtonInfo
    Caption         As String
    Key             As String
    Style           As ButtonStyleConstants
    Value           As ValueConstants
    ToolTipText     As String
    Enabled         As Boolean
    ChartIndex      As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TradingUI"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function CreateContractSelectionHelper( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pPreferredTickerGridRow As Long, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore) As ContractSelectionHelper
Const ProcName As String = "CreateContractSelectionHelper"
On Error GoTo Err

Set CreateContractSelectionHelper = New ContractSelectionHelper
CreateContractSelectionHelper.Initialise _
                        pContractSpec, _
                        pPreferredTickerGridRow, _
                        pPrimaryContractStore, _
                        pSecondaryContractStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SelectTickfiles( _
                ByRef UserCancelled As Boolean, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore, _
                Optional ByVal pTheme As ITheme) As TickfileSpecifiers
Const ProcName As String = "SelectTickfiles"
On Error GoTo Err

Static inHere As Boolean
Dim lTickfileOrganiser As fTickfileOrganiser

gLogger.Log "SelectTickfiles", ProcName, ModuleName, LogLevelMediumDetail

Assert Not inHere, "Already selecting tickfiles"
inHere = True

gLogger.Log "Create tickfile organiser form", ProcName, ModuleName, LogLevelMediumDetail
Set lTickfileOrganiser = New fTickfileOrganiser
lTickfileOrganiser.Initialise pTickfileStore, pPrimaryContractStore, pSecondaryContractStore

If Not pTheme Is Nothing Then lTickfileOrganiser.Theme = pTheme

gLogger.Log "Show tickfile organiser form", ProcName, ModuleName, LogLevelMediumDetail
lTickfileOrganiser.Show vbModal

gLogger.Log "Tickfile organiser form closed", ProcName, ModuleName, LogLevelMediumDetail

UserCancelled = lTickfileOrganiser.Cancelled

inHere = False
Set SelectTickfiles = lTickfileOrganiser.TickfileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




