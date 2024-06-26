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

Public Function CreateCurrencyConverter( _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pContractStore As IContractStore) As ICurrencyConverter
Const ProcName As String = "CreateCurrencyConverter"
On Error GoTo Err

Set CreateCurrencyConverter = GCurrencyUtils.CreateCurrencyConverter(pMarketDataManager, pContractStore)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptor( _
                ByVal CurrencyCode As String) As CurrencyDescriptor
Const ProcName As String = "GetCurrencyDescriptor"
On Error GoTo Err

GetCurrencyDescriptor = GCurrencyUtils.GetCurrencyDescriptor(CurrencyCode)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "GetCurrencyDescriptors"
On Error GoTo Err

GetCurrencyDescriptors = GCurrencyUtils.GetCurrencyDescriptors

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidCurrencyCode(ByVal CurrencyCode As String) As Boolean
Const ProcName As String = "IsValidCurrencyCode"
On Error GoTo Err

IsValidCurrencyCode = GCurrencyUtils.IsValidCurrencyCode(CurrencyCode)

Exit Function

Err:
GCurrency.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================





