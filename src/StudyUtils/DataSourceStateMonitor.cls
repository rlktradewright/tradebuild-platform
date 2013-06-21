VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceStateMonitor"
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

Implements StateChangeListener

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

Private Const ModuleName                            As String = "DataSourceStateMonitor"

'@================================================================================
' Member variables
'@================================================================================

Private mStudyBase                                  As TickDataStudyBase

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim lState As MarketDataSourceStates
lState = ev.State
If lState = MarketDataSourceStateReady Then
    Dim lDataSource As IMarketDataSource
    Set lDataSource = ev.Source
    mStudyBase.ReInitialise lDataSource.Contract, lDataSource.Session
    lDataSource.RemoveStateChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pStudyBase As TickDataStudyBase)
Set mStudyBase = pStudyBase
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





