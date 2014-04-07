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
    ' No chart has yet been started, or the chart has been cleared
    '
    '@/
    ChartStateBlank

    ''
    ' The chart has been created, but the Ticker has not yet become ready
    '
    '@/
    ChartStateCreated

    ''
    ' The Ticker is now ready and historical data has been requested
    '
    '@/
    ChartStateInitialised

    ''
    ' Historic data has been fetched and is being added to the chart
    '
    '@/
    ChartStateLoading
    
    ''
    ' All historic data has been added to the chart
    '
    '@/
    ChartStateLoaded
End Enum

Public Enum MultiChartChangeTypes
    MultiChartSelectionChanged
    MultiChartAdd
    MultiChartRemove
    MultiChartPeriodLengthChanged
End Enum

'@================================================================================
' Types
'@================================================================================

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
                ByVal pContractStore As IContractStore) As ContractSelectionHelper
Const ProcName As String = "CreateContractSelectionHelper"
On Error GoTo Err

Set CreateContractSelectionHelper = New ContractSelectionHelper
CreateContractSelectionHelper.Initialise pContractSpec, pContractStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SelectTickfiles( _
                ByRef UserCancelled As Boolean, _
                ByVal pTickfileStore As ITickfileStore, _
                ByVal pPrimaryContractStore As IContractStore, _
                Optional ByVal pSecondaryContractStore As IContractStore) As TickfileSpecifiers
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

gLogger.Log "Show tickfile organiser form", ProcName, ModuleName, LogLevelMediumDetail
lTickfileOrganiser.Show vbModal

gLogger.Log "Tickfile organiser form closed", ProcName, ModuleName, LogLevelMediumDetail

UserCancelled = lTickfileOrganiser.cancelled

inHere = False
Set SelectTickfiles = lTickfileOrganiser.TickfileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




