VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderNotifyReqs"
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "BracketOrderNotifyReqs"

'@================================================================================
' Member variables
'@================================================================================

Private mBracketOrderNotificationRequests           As New EnumerableCollection

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

Friend Function GetEnumerator(pBracketOrder) As Enumerator
Const ProcName As String = "GetEnumerator"
On Error GoTo Err

Dim lKey As String: lKey = GetObjectKey(pBracketOrder)
AssertArgument mBracketOrderNotificationRequests.Contains(lKey), "No notification requests for pBracketOrder"

Dim lRequests As EnumerableCollection
Set lRequests = mBracketOrderNotificationRequests.Item(lKey)

Set GetEnumerator = lRequests.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pStrategy As Object, _
                ByVal pResourceContext As ResourceContext)
Const ProcName As String = "Add"
On Error GoTo Err

Dim lKey As String: lKey = GetObjectKey(pBracketOrder)

Dim lRequests As EnumerableCollection
If mBracketOrderNotificationRequests.Contains(lKey) Then
    Set lRequests = mBracketOrderNotificationRequests.Item(lKey)
Else
    Set lRequests = New EnumerableCollection
    mBracketOrderNotificationRequests.Add lRequests, lKey
End If

Dim lNotificationRequest As New NotificationRequest
lNotificationRequest.Initialise pStrategy, pResourceContext
lRequests.Add lNotificationRequest

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




