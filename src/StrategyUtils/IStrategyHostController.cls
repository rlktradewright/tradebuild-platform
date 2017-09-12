VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStrategyHostController"
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

Private Const ModuleName                            As String = "IStrategyHostController"

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

Public Sub AddTimeframe( _
                ByVal pTimeframe As Timeframe)
End Sub

Public Sub ChartStudyValue(ByVal pStudy As IStudy, _
                ByVal pValueName As String, _
                ByVal pTimeframe As Timeframe)
End Sub

Public Sub ContractInvalid(ByVal pMessage As String)
End Sub

Public Sub Finish()
End Sub

Public Sub Initialise( _
                ByVal pStrategyRunner As StrategyRunner, _
                ByVal pModel As IStrategyHostModel, _
                ByVal pView As IStrategyHostView)
End Sub

Public Function GetDefaultParameters( _
                ByVal pStrategy As IStrategy, _
                ByVal pPositionManagementStrategyFactories As EnumerableCollection) As Parameters
End Function

Public Sub NotifyBracketOrderProfile(ByRef Value As BracketOrderProfile)
End Sub

Public Sub NotifyPosition(ByVal pPosition As Long)
End Sub

Public Sub NotifyReplayEvent(ev As NotificationEventData)
End Sub

Public Sub NotifyReplayCompleted()
End Sub

Public Sub NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
End Sub

Public Sub NotifyReplayStarted()
End Sub

Public Sub NotifyTickfileCompleted(ByVal pTickfile As TickfileSpecifier, ByVal pEventsPlayed As Long)
End Sub

Public Sub StartLiveProcessing(ByVal pSymbol As IContractSpecifier)
End Sub

Public Sub StartTickfileReplay(ByVal pTickFileSpecifiers As TickFileSpecifiers)
End Sub

Public Sub StopTickfileReplay()
End Sub

Public Sub TickerCreated(ByVal pTicker As Ticker)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

