VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Expiries"
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

Private Const ModuleName                            As String = "Expiries"

Private mExpiries()                                 As Date
Private mExpiriesNextIndex                          As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mExpiries(15) As Date
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

Public Property Get Count() As Long
Count = mExpiriesNextIndex
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pExpiry As String)
Const ProcName As String = "Add"
On Error GoTo Err

If mExpiriesNextIndex > UBound(mExpiries) Then
    ReDim Preserve mExpiries((UBound(mExpiries) + 1) * 2 - 1) As Date
End If

Dim lExpiry As Date: lExpiry = CDate(Left$(pExpiry, 4) & "-" & _
                                    Mid$(pExpiry, 5, 2) & "-" & _
                                    Right$(pExpiry, 2))
If mExpiriesNextIndex <> 0 Then AssertArgument lExpiry > mExpiries(mExpiriesNextIndex - 1), "Expiries must be added in ascending order"
mExpiries(mExpiriesNextIndex) = lExpiry

mExpiriesNextIndex = mExpiriesNextIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item(ByVal pIndex As Long) As Date
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Assert pIndex >= 0 And pIndex < mExpiriesNextIndex, _
        "pIndex invalid", _
        VBErrorCodes.VbErrSubscriptOutOfRange
Item = mExpiries(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================






