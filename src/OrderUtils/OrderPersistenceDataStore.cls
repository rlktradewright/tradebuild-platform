VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceDataStore"
Attribute VB_GlobalNameSpace = False
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

Implements IOrderPersistenceDataStore

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

Private Const ModuleName                            As String = "OrderPersistenceDataStore"

'@================================================================================
' Member variables
'@================================================================================

Private mRecoveryFilePath                           As String
Private mTextStreams                                As New EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IOrderPersistenceDataStore Interface Members
'@================================================================================

Private Sub IOrderPersistenceDataStore_Clear(ByVal pScopeName As String)
Const ProcName As String = "IOrderPersistenceDataStore_Clear"
On Error GoTo Err

If mTextStreams.Contains(pScopeName) Then mTextStreams.Remove pScopeName
getCurrentPersistenceFile pScopeName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IOrderPersistenceDataStore_Finish()
Const ProcName As String = "IOrderPersistenceDataStore_Finish"
On Error GoTo Err

Dim lTS As TextStream
For Each lTS In mTextStreams
    lTS.Close
Next

mTextStreams.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderPersistenceDataStore_GetEnumeratorForScope(ByVal pScopeName As String) As Enumerator
Const ProcName As String = "IOrderPersistenceDataStore_GetEnumeratorForScope"
On Error GoTo Err

Dim lOpEnumerator As New OrderPersistenceEnumerator
lOpEnumerator.Initialise getPreviousPersistenceFile(pScopeName)
Set IOrderPersistenceDataStore_GetEnumeratorForScope = lOpEnumerator.getEnumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IOrderPersistenceDataStore_SaveOrderInfo(ByVal pScopeName As String, ByVal pOrderInfo As String)
Const ProcName As String = "IOrderPersistenceDataStore_SaveOrderInfo"
On Error GoTo Err

getCurrentPersistenceFile(pScopeName).WriteLine pOrderInfo

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

Public Property Get RecoveryFilePath() As String
RecoveryFilePath = mRecoveryFilePath
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pRecoveryFilePath As String)
mRecoveryFilePath = pRecoveryFilePath
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getCurrentPersistenceFile(ByVal pScopeName As String) As TextStream
Const ProcName As String = "getCurrentPersistenceFile"
On Error GoTo Err

Dim fso As New FileSystemObject

If Not mTextStreams.Contains(pScopeName) Then
    Dim lTS As TextStream
    Set lTS = fso.CreateTextFile(getRecoveryFileName(pScopeName), True)
    mTextStreams.Add lTS, pScopeName
End If

Set getCurrentPersistenceFile = mTextStreams.Item(pScopeName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPreviousPersistenceFile(ByVal pScopeName As String) As TextStream
Const ProcName As String = "getPreviousPersistenceFile"
On Error GoTo Err

On Error Resume Next
Dim fso As New FileSystemObject
If fso.FileExists(getRecoveryFileName(pScopeName)) Then
    Set getPreviousPersistenceFile = fso.OpenTextFile(getRecoveryFileName(pScopeName), _
                            IOMode.ForReading, _
                            False)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getRecoveryFileName(ByVal pScopeName As String) As String
If pScopeName <> "" Then
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo(" & pScopeName & ").txt"
Else
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo.txt"
End If
End Function



