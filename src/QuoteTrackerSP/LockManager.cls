VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LockManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const LockIdBase As Long = 1024

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type LockEntry
    inUse               As Boolean
    key                 As String
    action              As String
    lockHolder          As LockConsumer
    waiters             As Queue
End Type

'================================================================================
' Member variables
'================================================================================

Private mLockTable() As LockEntry


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
ReDim mLockTable(10) As LockEntry

End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Function acquireLock( _
                ByVal key As String, _
                ByVal action As String, _
                ByVal locker As LockConsumer, _
                ByRef lockId As Long) As Boolean
Dim i As Long
Dim freeIndex As Long

freeIndex = -1
For i = 0 To UBound(mLockTable)
    If Not mLockTable(i).inUse Then freeIndex = i
    If mLockTable(i).key = key And mLockTable(i).action = action Then
        ' someone's already got a lock on this key/action
        If mLockTable(i).waiters Is Nothing Then
            Set mLockTable(i).waiters = New Queue
        End If
        mLockTable(i).waiters.addToEnd locker
        lockId = LockIdBase + i
        locker.waitingForLock lockId
        Exit Function
    End If
Next

If freeIndex = -1 Then
    freeIndex = UBound(mLockTable) + 1
    ReDim Preserve mLockTable(UBound(mLockTable) + 10) As LockEntry
End If

mLockTable(freeIndex).action = action
mLockTable(freeIndex).inUse = True
mLockTable(freeIndex).key = key
Set mLockTable(freeIndex).lockHolder = locker
lockId = LockIdBase + freeIndex

acquireLock = True

End Function

Friend Function acquireOrVerifyLock( _
                ByRef lockId As Long, _
                ByVal key As String, _
                ByVal action As String, _
                ByVal locker As LockConsumer) As Boolean

If lockId = 0 Then
    acquireOrVerifyLock = acquireLock(key, action, locker, lockId)
Else
    If Not verifyLock(lockId, locker) Then
        err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    "QTSP.LockManager::acquireOrVerifyLock", _
                    "Lock id " & lockId & " is not held by this object"
    End If
    acquireOrVerifyLock = True
End If

End Function

Friend Sub cancelAcquireLock( _
                ByVal key As String, _
                ByVal action As String, _
                ByVal locker As LockConsumer)
Dim i As Long

For i = 0 To UBound(mLockTable)
    If mLockTable(i).inUse And _
        mLockTable(i).key = key And _
        mLockTable(i).action = action _
    Then
        If mLockTable(i).lockHolder Is locker Then
            ' this locker already holds the lock so nothing to do
            Exit Sub
        End If

        If mLockTable(i).waiters Is Nothing Then
            err.Raise ErrorCodes.ErrIllegalArgumentException, _
                        "QTSP.LockManager::cancelAcquireLock", _
                        "Locker is not acquiring this lock"
        End If
        
        mLockTable(i).waiters.remove locker
        If mLockTable(i).waiters.size = 0 Then Set mLockTable(i).waiters = Nothing
        Exit Sub
    End If
Next


End Sub

Friend Sub releaseLock( _
                ByVal lockId As Long)

If mLockTable(lockId - LockIdBase).waiters Is Nothing Then
    releaseLockTableEntry lockId
Else
    Set mLockTable(lockId - LockIdBase).lockHolder = mLockTable(lockId - LockIdBase).waiters.removeFromFront
    If mLockTable(lockId - LockIdBase).waiters.size = 0 Then
        Set mLockTable(lockId - LockIdBase).waiters = Nothing
    End If
    mLockTable(lockId - LockIdBase).lockHolder.lockAcquired lockId
End If

End Sub

Friend Function verifyLock( _
                ByVal lockId As Long, _
                ByVal locker As LockConsumer) As Boolean
If mLockTable(lockId - LockIdBase).inUse And _
    mLockTable(lockId - LockIdBase).lockHolder Is locker _
Then
    verifyLock = True
End If
End Function

'================================================================================
' Helper Functions
'================================================================================

Friend Sub releaseLockTableEntry( _
                ByVal lockId As Long)
mLockTable(lockId - LockIdBase).action = ""
mLockTable(lockId - LockIdBase).inUse = False
mLockTable(lockId - LockIdBase).key = ""
Set mLockTable(lockId - LockIdBase).lockHolder = Nothing
Set mLockTable(lockId - LockIdBase).waiters = Nothing
End Sub


