VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccountValue"
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

Implements IAccountValue

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

Private Const ModuleName                            As String = "AccountValue"

'@================================================================================
' Member variables
'@================================================================================

Private mAccountName                                As String
Private mCurrencyCode                               As String
Private mKey                                        As String
Private mValue                                      As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IAccountValue Interface Members
'@================================================================================

Private Property Get IAccountValue_AccountName() As String
IAccountValue_AccountName = mAccountName
End Property

Private Property Get IAccountValue_CurrencyCode() As String
IAccountValue_CurrencyCode = mCurrencyCode
End Property

Private Property Get IAccountValue_Key() As String
IAccountValue_Key = mKey
End Property

Private Property Get IAccountValue_Value() As String
IAccountValue_Value = mValue
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pAccountName As String, _
                ByVal pCurrencyCode As String, _
                ByVal pKey As String, _
                ByVal pValue As String)
mAccountName = pAccountName
mCurrencyCode = pCurrencyCode
mKey = pKey
mValue = pValue
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




