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

Public Enum PermittedServiceProviders
    SPRealtimeData = 1
    SPPrimaryContractData = 2
    SPSecondaryContractData = 4
    SPHistoricalDataInput = 8
    SPHistoricalDataOutput = &H10
    SPBroker = &H20
    SPTickfileInput = &H40
    SPTickfileOutput = &H80
End Enum

Public Enum ChartStates
    ''
    ' The chart is ready to have studies added to it
    '
    '@/
    ChartStateInitialised
    
    ''
    ' All historic data has been added to the chart
    '
    '@/
    ChartStateLoaded
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

Public Function SelectTickfiles() As TickfileSelection
Static inHere As Boolean

If inHere Then
    ' this is a re-entrant call
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "SelectTickfiles", _
            "Already selecting tickfiles"
End If

Set SelectTickfiles = New TickfileSelection
SelectTickfiles.SelectTickfiles
End Function

'@================================================================================
' Helper Functions
'@================================================================================


