VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandleAllocator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Type HandlesEntry
    Next            As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "HandleAllocator"

'@================================================================================
' Member variables
'@================================================================================

Private mHandles()                                  As HandlesEntry
Private mHandlesIndex                               As Long
Private mNextFreeHandlesIndex                       As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

ReDim mHandles(7) As HandlesEntry
mNextFreeHandlesIndex = NullIndex

Exit Sub

Err:
GMktData.HandleUnexpectedError ProcName, ModuleName
End Sub

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

Friend Function AllocateHandle() As Long
If mNextFreeHandlesIndex <> NullIndex Then
    AllocateHandle = mNextFreeHandlesIndex
    mNextFreeHandlesIndex = mHandles(mNextFreeHandlesIndex).Next
Else
    If mHandlesIndex > UBound(mHandles) Then ReDim Preserve mHandles(2 * (UBound(mHandles) + 1) - 1) As HandlesEntry
    AllocateHandle = mHandlesIndex
    mHandlesIndex = mHandlesIndex + 1
End If
End Function

Friend Sub ReleaseHandle(ByVal pHandle As Long)
' we maintain the free list so that lower indexes are
' reallocated first
If mNextFreeHandlesIndex = NullIndex Then
    mHandles(pHandle).Next = NullIndex
    mNextFreeHandlesIndex = pHandle
Else
    Dim lCurrFreeIndex As Long
    lCurrFreeIndex = mNextFreeHandlesIndex
    
    Dim lPrevFreeIndex As Long
    Do While pHandle > lCurrFreeIndex And lCurrFreeIndex <> NullIndex
        lPrevFreeIndex = lCurrFreeIndex
        lCurrFreeIndex = mHandles(lCurrFreeIndex).Next
    Loop
    mHandles(lPrevFreeIndex).Next = pHandle
    mHandles(pHandle).Next = lCurrFreeIndex
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




