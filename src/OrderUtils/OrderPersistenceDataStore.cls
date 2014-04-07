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

Implements Enumerable
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

Private Const OrderInfoSession                      As String = "SESSION"

'@================================================================================
' Member variables
'@================================================================================

Private mRecoveryFilePath                           As String
Private mScopeName                                  As String
Private mOutputTS                                       As TextStream

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"
On Error GoTo Err

Set Enumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_NextItem"
On Error GoTo Err

Dim lTS As TextStream
Set lTS = cookie

Dim lLine As String
If lTS Is Nothing Then
    Enumerable_NextItem = False
ElseIf getNextLine(lTS, lLine) Then
    Item = lLine
    Enumerable_NextItem = True
Else
    Enumerable_NextItem = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Enumerable_Remove(cookie As Variant)
Const ProcName As String = "Enumerable_Remove"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Remove not supported"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_Skip"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not supported"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' IOrderPersistenceDataStore Interface Members
'@================================================================================

Private Sub IOrderPersistenceDataStore_Clear()
Const ProcName As String = "IOrderPersistenceDataStore_Clear"
On Error GoTo Err

Set mOutputTS = Nothing
getCurrentPersistenceFile

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IOrderPersistenceDataStore_GetEnumeratorForScope(ByVal pScopeName As String) As Enumerator
Const ProcName As String = "IOrderPersistenceDataStore_GetEnumeratorForScope"
On Error GoTo Err

mScopeName = pScopeName
Set IOrderPersistenceDataStore_GetEnumeratorForScope = getEnumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IOrderPersistenceDataStore_SaveOrderInfo(ByVal pOrderInfo As String)
Const ProcName As String = "IOrderPersistenceDataStore_SaveOrderInfo"
On Error GoTo Err

getCurrentPersistenceFile.WriteLine pOrderInfo

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

Private Function getCurrentPersistenceFile() As TextStream
Const ProcName As String = "getCurrentPersistenceFile"
On Error GoTo Err

Dim fso As New FileSystemObject

If mOutputTS Is Nothing Then Set mOutputTS = fso.CreateTextFile(getRecoveryFileName(), True)
Set getCurrentPersistenceFile = mOutputTS

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Assert mScopeName <> "", "Scopename has not been set"

Dim lTS As TextStream
Set lTS = getPreviousPersistenceFile
Set getEnumerator = CreateEnumerator(Me, lTS)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getNextLine(ByVal pTs As TextStream, ByRef result As String) As Boolean
Const ProcName As String = "getNextLine"
On Error GoTo Err

Do While Not pTs.AtEndOfStream
    result = pTs.ReadLine
    If result <> "" And Left$(result, 2) <> "//" Then
        getNextLine = True
        Exit Do
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPreviousPersistenceFile() As TextStream
Const ProcName As String = "getPreviousPersistenceFile"
On Error GoTo Err

On Error Resume Next
Dim fso As New FileSystemObject
Set getPreviousPersistenceFile = fso.OpenTextFile(getRecoveryFileName(), _
                        IOMode.ForReading, _
                        False)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getRecoveryFileName() As String
If mScopeName <> "" Then
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo(" & mScopeName & ").txt"
Else
    getRecoveryFileName = mRecoveryFilePath & "\OrderInfo.txt"
End If
End Function



