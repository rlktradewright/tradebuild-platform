VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TradingDBApi"
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TradingDBApi"

'@================================================================================
' Member variables
'@================================================================================

Private mTradingDO                                  As New TradingDO

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

Public Function CreateTradingDBClient( _
                ByVal pDatabaseType As DatabaseTypes, _
                ByVal pServer As String, _
                ByVal pDatabaseName As String, _
                Optional ByVal pUsername As String, _
                Optional ByVal pPassword As String, _
                Optional ByVal pUseSynchronousReads As Boolean, _
                Optional ByVal pUseSynchronousWrites As Boolean) As DBClient
Const ProcName As String = "CreateTradingDBClient"
On Error GoTo Err

Dim lDBClient As New DBClient
lDBClient.Initialise CreateTradingDBFuture(CreateConnectionParams( _
                                        pDatabaseType, _
                                        pServer, _
                                        pDatabaseName, _
                                        pUsername, _
                                        pPassword)), _
                    pUseSynchronousReads, _
                    pUseSynchronousWrites

Set CreateTradingDBClient = lDBClient
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DatabaseTypeFromString( _
                ByVal Value As String) As DatabaseTypes
Const ProcName As String = "DatabaseTypeFromString"
On Error GoTo Err

DatabaseTypeFromString = mTradingDO.DatabaseTypeFromString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DatabaseTypeToString( _
                ByVal Value As DatabaseTypes) As String
Const ProcName As String = "DatabaseTypeToString"
On Error GoTo Err

DatabaseTypeToString = mTradingDO.DatabaseTypeToString(Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




