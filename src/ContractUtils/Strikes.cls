VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Strikes"
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

Private Const ModuleName                            As String = "Strikes"

'@================================================================================
' Member variables
'@================================================================================

Private mStrikes()                                  As Double
Private mStrikesNextIndex                           As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStrikes(15) As Double
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
Count = mStrikesNextIndex
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pStrike As Double)
Const ProcName As String = "Add"
On Error GoTo Err

If mStrikesNextIndex <> 0 Then AssertArgument pStrike > mStrikes(mStrikesNextIndex - 1), "Strikes must be added in ascending order"
If mStrikesNextIndex > UBound(mStrikes) Then
    ReDim Preserve mStrikes((UBound(mStrikes) + 1) * 2 - 1) As Double
End If
mStrikes(mStrikesNextIndex) = pStrike
mStrikesNextIndex = mStrikesNextIndex + 1

Exit Sub

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item(ByVal pIndex As Long) As Double
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Assert pIndex >= 0 And pIndex < mStrikesNextIndex, _
        "pIndex invalid", _
        VBErrorCodes.VbErrSubscriptOutOfRange

Item = mStrikes(pIndex)

Exit Function

Err:
GContracts.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




