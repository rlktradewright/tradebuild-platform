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
    ' No chart has yet been started, or a chart has been cleared
    '
    '@/
    ChartStateBlank

    ''
    ' The chart has been created, but the ticker has not yet become ready
    '
    '@/
    ChartStateCreated

    ''
    ' The ticker is now ready and historical data has been requested
    '
    '@/
    ChartStateInitialised

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

Private Const ProjectName                   As String = "TradeBuildUI26"
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
                ByRef userCancelled As Boolean) As TickfileSpecifiers
Static inHere As Boolean

Dim lTickfileOrganiser As fTickfileOrganiser
Dim lLogger As Logger


Dim failpoint As Long
On Error GoTo Err

Set lLogger = GetLogger("diag.tradebuild.tradebuildui")
lLogger.Log LogLevelMediumDetail, "SelectTickfiles"

If inHere Then
    ' this is a re-entrant call
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SelectTickfiles", _
            "Already selecting tickfiles"
End If

lLogger.Log LogLevelMediumDetail, "Create tickfile organiser form"
Set lTickfileOrganiser = New fTickfileOrganiser

lLogger.Log LogLevelMediumDetail, "Show tickfile organiser form"
lTickfileOrganiser.Show vbModal

lLogger.Log LogLevelMediumDetail, "Tickfile organiser form closed"

userCancelled = lTickfileOrganiser.cancelled

Set SelectTickfiles = lTickfileOrganiser.TickfileSpecifiers

Exit Function

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SelectTickfiles" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription

End Function

'@================================================================================
' Helper Functions
'@================================================================================


