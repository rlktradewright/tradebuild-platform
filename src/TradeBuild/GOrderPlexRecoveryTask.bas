Attribute VB_Name = "GOrderPlexRecoveryTask"
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

Public Enum OpRecoveryActions
    ActionSyncEntryOrder = 1
    ActionSyncStopOrder
    ActionSyncTargetOrder
    ActionSyncCloseoutOrder
    ActionCompleteEntryOrder
    ActionRecoveryCompletion
End Enum

Public Enum OpRecoveryConditions
    CondTickerSet = 1
    CondEntryOrderSet = 2
    CondStopOrderSet = 4
    CondTargetOrderSet = 8
    CondCloseoutOrderSet = &H10
    CondEntryOrderRecovered = &H20
    CondStopOrderRecovered = &H40
    CondTargetOrderRecovered = &H80
    CondCloseoutOrderRecovered = &H100
End Enum

Public Enum OpRecoveryStates
    Started = 1
    OrderPlexRecreated
    Finished
End Enum

Public Enum OpRecoveryStimuli
    StimSetTicker
    StimRecreated
    StimRecoverEntryOrder
    StimRecoverStopOrder
    StimRecoverTargetOrder
    StimRecoverCloseoutOrder
    StimSetOpState
    
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GOrderPlexRecoveryTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTableBuilder As StateTableBuilder

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

Public Property Get TableBuilder() As StateTableBuilder
If mTableBuilder Is Nothing Then
    Set mTableBuilder = New StateTableBuilder
    buildStateTable
End If
Set TableBuilder = mTableBuilder
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub buildStateTable()

'=======================================================================
'                       State:      Started
'=======================================================================

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.Started, _
            OpRecoveryStimuli.StimRecreated, _
            OpRecoveryConditions.CondEntryOrderSet, _
            SpecialConditions.NoConditions, _
            OpRecoveryStates.OrderPlexRecreated, _
            SpecialActions.NoAction

'=======================================================================
'                       State:      OrderPlexRecreated
'=======================================================================

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverEntryOrder, _
            OpRecoveryConditions.CondEntryOrderSet, _
            OpRecoveryConditions.CondEntryOrderRecovered + _
                    OpRecoveryConditions.CondStopOrderSet + _
                    OpRecoveryConditions.CondTargetOrderSet, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionSyncEntryOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverEntryOrder, _
            OpRecoveryConditions.CondEntryOrderSet, _
            OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryActions.ActionSyncEntryOrder


mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverStopOrder, _
            OpRecoveryConditions.CondStopOrderSet, _
            OpRecoveryConditions.CondStopOrderRecovered + _
                OpRecoveryConditions.CondEntryOrderRecovered + _
                OpRecoveryConditions.CondTargetOrderSet, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionCompleteEntryOrder, _
            OpRecoveryActions.ActionSyncStopOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverStopOrder, _
            OpRecoveryConditions.CondStopOrderSet, _
            OpRecoveryConditions.CondStopOrderRecovered + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryActions.ActionCompleteEntryOrder, _
            OpRecoveryActions.ActionSyncStopOrder

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverStopOrder, _
            OpRecoveryConditions.CondStopOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryConditions.CondStopOrderRecovered + _
                OpRecoveryConditions.CondTargetOrderSet, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionSyncStopOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverStopOrder, _
            OpRecoveryConditions.CondStopOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered + _
                OpRecoveryConditions.CondTargetOrderRecovered, _
            OpRecoveryConditions.CondStopOrderRecovered, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionSyncStopOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverStopOrder, _
            OpRecoveryConditions.CondStopOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryConditions.CondStopOrderRecovered, _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryActions.ActionSyncStopOrder


mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverTargetOrder, _
            OpRecoveryConditions.CondTargetOrderSet, _
            OpRecoveryConditions.CondTargetOrderRecovered + _
                OpRecoveryConditions.CondEntryOrderRecovered + _
                OpRecoveryConditions.CondStopOrderSet, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionCompleteEntryOrder, _
            OpRecoveryActions.ActionSyncTargetOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverTargetOrder, _
            OpRecoveryConditions.CondTargetOrderSet, _
            OpRecoveryConditions.CondTargetOrderRecovered + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryActions.ActionCompleteEntryOrder, _
            OpRecoveryActions.ActionSyncTargetOrder

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverTargetOrder, _
            OpRecoveryConditions.CondTargetOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryConditions.CondTargetOrderRecovered + _
                OpRecoveryConditions.CondStopOrderSet, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionSyncTargetOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverTargetOrder, _
            OpRecoveryConditions.CondTargetOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered + _
                OpRecoveryConditions.CondStopOrderRecovered, _
            OpRecoveryConditions.CondTargetOrderRecovered, _
            OpRecoveryStates.Finished, _
            OpRecoveryActions.ActionSyncTargetOrder, _
            OpRecoveryActions.ActionRecoveryCompletion

mTableBuilder.addStateTableEntry _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryStimuli.StimRecoverTargetOrder, _
            OpRecoveryConditions.CondTargetOrderSet + _
                OpRecoveryConditions.CondEntryOrderRecovered, _
            OpRecoveryConditions.CondTargetOrderRecovered, _
            OpRecoveryStates.OrderPlexRecreated, _
            OpRecoveryActions.ActionSyncTargetOrder

mTableBuilder.stateTableComplete
End Sub

