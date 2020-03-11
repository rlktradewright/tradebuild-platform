VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CloseoutSpecifier"
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

Private Const ModuleName                            As String = "CloseoutSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderType                                  As OrderTypes
Private mLimitPriceSpec                             As PriceSpecifier
Private mTriggerPriceSpec                           As PriceSpecifier
Private mOutsideRTH                                 As Boolean
Private mForceCloseAfterSecs                        As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONable Interface Members
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

Public Property Let OrderType(ByVal Value As OrderTypes)
mOrderType = Value
End Property

Public Property Get OrderType() As OrderTypes
OrderType = mOrderType
End Property

Public Property Let LimitPriceSpec(ByVal Value As PriceSpecifier)
Set mLimitPriceSpec = Value
End Property

Public Property Get LimitPriceSpec() As PriceSpecifier
Set LimitPriceSpec = mLimitPriceSpec
End Property

Public Property Let TriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mTriggerPriceSpec = Value
End Property

Public Property Get TriggerPriceSpec() As PriceSpecifier
Set TriggerPriceSpec = mTriggerPriceSpec
End Property

Public Property Let OutsideRTH(ByVal Value As Boolean)
mOutsideRTH = Value
End Property

Public Property Get OutsideRTH() As Boolean
OutsideRTH = mOutsideRTH
End Property

Public Property Let ForceCloseAfterSecs(ByVal Value As Long)
mForceCloseAfterSecs = Value
End Property

Public Property Get ForceCloseAfterSecs() As Long
ForceCloseAfterSecs = mForceCloseAfterSecs
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "OrderType", mOrderType, OrderTypeMarket
jb.EncodeNameAndValue "OutsideRTH", mOutsideRTH, False
jb.EncodeNameAndValue "ForceCloseAfterSecs", mForceCloseAfterSecs, 0
If Not gIsNullPriceSpecifier(mLimitPriceSpec) Then _
    jb.EncodeNameAndValue "LimitPriceSpec", mLimitPriceSpec, Nothing
If Not gIsNullPriceSpecifier(mTriggerPriceSpec) Then _
    jb.EncodeNameAndValue "TriggerPriceSpec", mTriggerPriceSpec, Nothing

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




