VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigChangeMonitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ConfigChangeMonitor"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mConfigStore                     As ConfigurationStore
Attribute mConfigStore.VB_VarHelpID = -1

Private WithEvents mTLI                             As TimerListItem
Attribute mTLI.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mConfigStore Event Handlers
'@================================================================================

Private Sub mConfigStore_Change(ev As TWUtilities40.ChangeEventData)
Const ProcName As String = "mConfigStore_Change"
On Error GoTo Err

If ev.ChangeType = ConfigChangeTypes.ConfigDirty Then
    If mTLI Is Nothing Then Set mTLI = GetGlobalTimerList.Add(Empty, 250, ExpiryTimeUnitMilliseconds)
ElseIf ev.ChangeType = ConfigChangeTypes.ConfigClean Then
    If Not mTLI Is Nothing Then
        mTLI.Cancel
        Set mTLI = Nothing
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTLI Event Handlers
'@================================================================================

Private Sub mTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mTLI_StateChange"
On Error GoTo Err

If Not ev.State = TimerListItemStates.TimerListItemStateExpired Then Exit Sub

LogMessage "Saving configuration"
mConfigStore.Save
Set mTLI = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mTLI Is Nothing Then
    mTLI.Cancel
    LogMessage "Saving configuration"
    mConfigStore.Save
    Set mTLI = Nothing
End If
Set mConfigStore = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise(ByVal pConfigStore As ConfigurationStore)
Set mConfigStore = pConfigStore
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




