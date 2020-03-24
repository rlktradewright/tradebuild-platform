VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RolloverSpecification"
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

Private Const ModuleName                            As String = "RolloverSpecification"

'@================================================================================
' Member variables
'@================================================================================

Private mDays                                       As Long
Private mTime                                       As Date

Private mStrikeMode                                 As RolloverStrikeModes
Private mStrikeValue                                As Long

Private mUnderlyingExchangeName                     As String

Private mCloseOrderType                             As OrderTypes
Private mCloseLimitPriceSpec                        As PriceSpecifier
Private mCloseTriggerPriceSpec                      As PriceSpecifier
Private mCloseTimeoutSecs                           As Long

Private mEntryOrderType                             As OrderTypes
Private mEntryLimitPriceSpec                        As PriceSpecifier
Private mEntryTriggerPriceSpec                      As PriceSpecifier
Private mEntryTimeoutSecs                           As Long

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

Public Property Let CloseOrderType(ByVal Value As OrderTypes)
mCloseOrderType = Value
End Property

Public Property Get CloseOrderType() As OrderTypes
CloseOrderType = mCloseOrderType
End Property

Public Property Get CloseLimitPriceSpec() As PriceSpecifier
Set CloseLimitPriceSpec = mCloseLimitPriceSpec
End Property

Public Property Let CloseLimitPriceSpec(ByVal Value As PriceSpecifier)
Set mCloseLimitPriceSpec = Value
End Property

Public Property Get CloseTimeoutSecs() As Long
CloseTimeoutSecs = mCloseTimeoutSecs
End Property

Public Property Let CloseTimeoutSecs(ByVal Value As Long)
mCloseTimeoutSecs = Value
End Property

Public Property Get CloseTriggerPriceSpec() As PriceSpecifier
Set CloseTriggerPriceSpec = mCloseTriggerPriceSpec
End Property

Public Property Let CloseTriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mCloseTriggerPriceSpec = Value
End Property

Public Property Let Days(ByVal Value As Long)
mDays = Value
End Property

Public Property Get Days() As Long
Days = mDays
End Property

Public Property Let EntryOrderType(ByVal Value As OrderTypes)
mEntryOrderType = Value
End Property

Public Property Get EntryOrderType() As OrderTypes
EntryOrderType = mEntryOrderType
End Property

Public Property Get EntryLimitPriceSpec() As PriceSpecifier
Set EntryLimitPriceSpec = mEntryLimitPriceSpec
End Property

Public Property Let EntryLimitPriceSpec(ByVal Value As PriceSpecifier)
Set mEntryLimitPriceSpec = Value
End Property

Public Property Get EntryTimeoutSecs() As Long
EntryTimeoutSecs = mEntryTimeoutSecs
End Property

Public Property Let EntryTimeoutSecs(ByVal Value As Long)
mEntryTimeoutSecs = Value
End Property

Public Property Get EntryTriggerPriceSpec() As PriceSpecifier
Set EntryTriggerPriceSpec = mEntryTriggerPriceSpec
End Property

Public Property Let EntryTriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mEntryTriggerPriceSpec = Value
End Property

Public Property Get StrikeMode() As RolloverStrikeModes
StrikeMode = mStrikeMode
End Property

Public Property Let StrikeMode(ByVal Value As RolloverStrikeModes)
mStrikeMode = Value
End Property

Public Property Get StrikeValue() As Long
StrikeValue = mStrikeValue
End Property

Public Property Let StrikeValue(ByVal Value As Long)
mStrikeValue = Value
End Property

Public Property Let Time(ByVal Value As Date)
mTime = Value
End Property

Public Property Get Time() As Date
Time = mTime
End Property

Public Property Get UnderlyingExchangeName() As String
UnderlyingExchangeName = mUnderlyingExchangeName
End Property

Public Property Let UnderlyingExchangeName(ByVal Value As String)
mUnderlyingExchangeName = Value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier)
mDays = pDays
mTime = pTime
mCloseOrderType = pCloseOrderType
mCloseTimeoutSecs = pCloseTimeoutSecs
Set mCloseLimitPriceSpec = pCloseLimitPriceSpec
Set mCloseTriggerPriceSpec = pCloseTriggerPriceSpec
mEntryOrderType = pEntryOrderType
mEntryTimeoutSecs = pEntryTimeoutSecs
Set mEntryLimitPriceSpec = pEntryLimitPriceSpec
Set mEntryTriggerPriceSpec = pEntryTriggerPriceSpec
End Sub

Friend Sub InitialiseForOption( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pStrikeMode As RolloverStrikeModes, _
                ByVal pStrikeValue As Long, _
                ByVal pUnderlyingExchangeName As String, _
                ByVal pCloseOrderType As OrderTypes, _
                ByVal pCloseTimeoutSecs As Long, _
                ByVal pCloseLimitPriceSpec As PriceSpecifier, _
                ByVal pCloseTriggerPriceSpec As PriceSpecifier, _
                ByVal pEntryOrderType As OrderTypes, _
                ByVal pEntryTimeoutSecs As Long, _
                ByVal pEntryLimitPriceSpec As PriceSpecifier, _
                ByVal pEntryTriggerPriceSpec As PriceSpecifier)
mDays = pDays
mTime = pTime
mStrikeMode = pStrikeMode
mStrikeValue = pStrikeValue
mUnderlyingExchangeName = pUnderlyingExchangeName
mCloseOrderType = pCloseOrderType
mCloseTimeoutSecs = pCloseTimeoutSecs
Set mCloseLimitPriceSpec = pCloseLimitPriceSpec
Set mCloseTriggerPriceSpec = pCloseTriggerPriceSpec
mEntryOrderType = pEntryOrderType
mEntryTimeoutSecs = pEntryTimeoutSecs
Set mEntryLimitPriceSpec = pEntryLimitPriceSpec
Set mEntryTriggerPriceSpec = pEntryTriggerPriceSpec
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "Days", mDays, 0

jb.EncodeNameAndValue "Time", CDbl(mTime), 0

jb.EncodeNameAndValue "StrikeMode", mStrikeMode, RolloverStrikeModeNone
jb.EncodeNameAndValue "StrikeValue", mStrikeValue, 0

jb.EncodeNameAndValue "UnderlyingExchangeName", mUnderlyingExchangeName, ""

jb.EncodeNameAndValue "CloseOrderType", mCloseOrderType, OrderTypeNone
jb.EncodeNameAndValue "CloseTimeoutSecs", mCloseTimeoutSecs, 0

If Not gIsNullPriceSpecifier(mCloseLimitPriceSpec) Then _
    jb.EncodeNameAndValue "CloseLimitPriceSpec", mCloseLimitPriceSpec, Nothing
If Not gIsNullPriceSpecifier(mCloseTriggerPriceSpec) Then _
    jb.EncodeNameAndValue "CloseTriggerPriceSpec", mCloseTriggerPriceSpec, Nothing

jb.EncodeNameAndValue "EntryOrderType", mEntryOrderType, OrderTypeNone
jb.EncodeNameAndValue "EntryTimeoutSecs", mEntryTimeoutSecs, 0

If Not gIsNullPriceSpecifier(mEntryLimitPriceSpec) Then _
    jb.EncodeNameAndValue "EntryLimitPriceSpec", mEntryLimitPriceSpec, Nothing
If Not gIsNullPriceSpecifier(mEntryTriggerPriceSpec) Then _
    jb.EncodeNameAndValue "EntryTriggerPriceSpec", mEntryTriggerPriceSpec, Nothing

jb.EndObject

ToJSON = jb.JSONString
End Function




