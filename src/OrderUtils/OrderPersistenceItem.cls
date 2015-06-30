VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceItem"
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

Private Const ModuleName                            As String = "OrderPersistenceItem"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestamp                                  As Date
Private mPositionManagerName                        As String
Private mData                                       As Variant
Private mItemType                                   As String
Private mProgId                                     As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONAble Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Data(ByVal Value As Variant)
gSetVariant mData, Value
End Property

Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

Public Property Let ItemType(ByVal Value As String)
mItemType = Value
End Property

Public Property Get ItemType() As String
ItemType = mItemType
End Property

Public Property Let PositionManagerName(ByVal Value As String)
mPositionManagerName = Value
End Property

Public Property Get PositionManagerName() As String
PositionManagerName = mPositionManagerName
End Property

Public Property Let ProgId(ByVal Value As String)
mProgId = Value
End Property

Public Property Get ProgId() As String
ProgId = mProgId
End Property

Public Property Let Timestamp(ByVal Value As Date)
mTimestamp = Value
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "PositionManagerName"
jb.EncodeValue mPositionManagerName

jb.EncodeName "Timestamp"
jb.EncodeValue mTimestamp

jb.EncodeName "ItemType"
jb.EncodeValue mItemType

jb.EncodeName "ProgId"
jb.EncodeValue mProgId

jb.EncodeName "Data"
jb.EncodeValue mData

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


