VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScanResults"
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

Implements IScanResults

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

Private Const ModuleName                            As String = "ScanResults"

'@================================================================================
' Member variables
'@================================================================================

Private mResults                                    As EnumerableCollection

Private mEnumerator                                 As Enumerator

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mResults = New EnumerableCollection
End Sub

'@================================================================================
' IScanResults Interface Members
'@================================================================================

Private Function IScanResults_Count() As Long
IScanResults_Count = mResults.Count
End Function

Private Property Get IScanResults_Enumerator() As Enumerator
Set IScanResults_Enumerator = mResults.Enumerator
End Property

Private Function IScanResults_Item(ByVal pIndex As Long) As IScanResult
Set IScanResults_Item = mResults.Item(CStr(pIndex))
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Function Add(ByVal pResult As IScanResult) As IScanResult
Const ProcName As String = "Add"
On Error GoTo Err

mResults.Add pResult, CStr(pResult.Rank)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


