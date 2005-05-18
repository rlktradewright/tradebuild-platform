VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServiceProviderError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables and constants
'================================================================================

Private mTimestamp As Date
Private mErrorCode As Long
Private mServiceProviderName As String
Private mMessage As String


'================================================================================
' Enums
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Let errorCode(ByVal value As Long)
mErrorCode = value
End Property

Public Property Get errorCode() As Long
errorCode = mErrorCode
End Property

Friend Property Let message(ByVal value As String)
mMessage = value
End Property

Public Property Get message() As String
message = mMessage
End Property

Friend Property Let serviceProviderName(ByVal value As String)
mServiceProviderName = value
End Property

Public Property Get serviceProviderName() As String
serviceProviderName = mServiceProviderName
End Property

Friend Property Let timestamp(ByVal value As Date)
mTimestamp = value
End Property

Public Property Get timestamp() As Date
timestamp = mTimestamp
End Property



'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


