VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ROD"
Attribute VB_GlobalNameSpace = False
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

Implements IJSONable

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

Private Const ModuleName                            As String = "ROD"

'@================================================================================
' Member variables
'@================================================================================

Private mProperties                                 As New Dictionary

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
Const ProcName As String = "IJSONable_ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "Properties"
jb.EncodeValue mProperties

jb.EndObject

IJSONAble_ToJSON = jb.JSONString

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Properties(ByVal Value As Dictionary)
Set mProperties = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Clear()
Set mProperties = New Dictionary
End Sub

Friend Sub SetProperty( _
                ByVal pName As String, _
                ByVal pValue As Variant)
Const ProcName As String = "SetProperty"
On Error GoTo Err

If mProperties.Exists(pName) Then mProperties.Remove pName
mProperties.Add pName, pValue

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UpdateHost(ByVal pHost As Object)
Const ProcName As String = "UpdateHost"
On Error GoTo Err

Dim lKeys() As Variant
lKeys = mProperties.Keys

Dim lKey As Variant
For Each lKey In lKeys
    CallByName pHost, lKey, VbLet, mProperties.Item(lKey)
Next

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




