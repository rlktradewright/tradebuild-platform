VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyUtils"
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

Public Type CurrencyDescriptor
    Code        As String
    Description         As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "CurrencyUtils"

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

Public Function GetCurrencyDescriptor( _
                ByVal CurrencyCode As String) As CurrencyDescriptor
Const ProcName As String = "GetCurrencyDescriptor"
On Error GoTo Err

GetCurrencyDescriptor = gGetCurrencyDescriptor(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "GetCurrencyDescriptors"
On Error GoTo Err

GetCurrencyDescriptors = gGetCurrencyDescriptors

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidCurrencyCode(ByVal CurrencyCode As String) As Boolean
Const ProcName As String = "IsValidCurrencyCode"
On Error GoTo Err

IsValidCurrencyCode = gIsValidCurrencyCode(CurrencyCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





