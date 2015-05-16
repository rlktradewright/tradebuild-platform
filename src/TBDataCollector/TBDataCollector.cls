VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TBDataCollector"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const ModuleName                    As String = "GlobalClass"

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

Public Function CreateDataCollector( _
                ByVal pConfigFile As ConfigurationStore, _
                ByVal pConfigName As String, _
                ByVal pStartTimeDescriptor As String, _
                ByVal pEndTimeDescriptor As String, _
                ByVal pExitTimeDescriptor As String, _
                Optional ByVal pSaveIntervalSeconds As Long = 5, _
                Optional ByVal pBarOutputMonitor As IBarOutputMonitor, _
                Optional ByVal pTickfileOutputMonitor As ITickfileOutputMonitor) As DataCollector
Const ProcName As String = "CreateDataCollector"
On Error GoTo Err

Set CreateDataCollector = New DataCollector
CreateDataCollector.Initialise pConfigFile, _
                            pConfigName, _
                            pStartTimeDescriptor, _
                            pEndTimeDescriptor, _
                            pExitTimeDescriptor, _
                            pSaveIntervalSeconds, _
                            pBarOutputMonitor, _
                            pTickfileOutputMonitor

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


