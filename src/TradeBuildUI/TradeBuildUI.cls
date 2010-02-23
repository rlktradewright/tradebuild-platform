VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradeBuildUI"
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
    MultiChartTimeframeChanged
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "TradeBuildUI"

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

Public Function SelectTickfiles( _
                ByRef UserCancelled As Boolean) As TickfileSpecifiers
Static inHere As Boolean

Dim lTickfileOrganiser As fTickfileOrganiser
Dim lLogger As Logger

Const ProcName As String = "SelectTickfiles"
Dim failpoint As String
On Error GoTo Err

Set lLogger = GetLogger("diag.tradebuild.tradebuildui")
lLogger.Log LogLevelMediumDetail, "SelectTickfiles"

If inHere Then
    ' this is a re-entrant call
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already selecting tickfiles"
End If

lLogger.Log LogLevelMediumDetail, "Create tickfile organiser form"
Set lTickfileOrganiser = New fTickfileOrganiser

lLogger.Log LogLevelMediumDetail, "Show tickfile organiser form"
lTickfileOrganiser.Show vbModal

lLogger.Log LogLevelMediumDetail, "Tickfile organiser form closed"

UserCancelled = lTickfileOrganiser.cancelled

Set SelectTickfiles = lTickfileOrganiser.TickfileSpecifiers

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

Public Function StartTickersWithUserSelection( _
                ByVal pWorkspace As Workspace, _
                ByVal pTickerOptions As TickerOptions, _
                ByVal contractSpec As contractSpecifier, _
                ByVal pContractSelector As IContractSelector, _
                Optional ByVal startIntervalMillisecs As Long = 50) As TickersStarter
Const ProcName As String = "StartTickersWithUserSelection"
Dim failpoint As String
On Error GoTo Err

Set StartTickersWithUserSelection = New TickersStarter
StartTickersWithUserSelection.Initialise pWorkspace, _
                                        pTickerOptions, _
                                        contractSpec, _
                                        pContractSelector, _
                                        startIntervalMillisecs

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


