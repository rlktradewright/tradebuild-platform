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

Private mInitialStrikeSelectionMode                 As OptionStrikeSelectionModes
Private mInitialStrikeParameter                     As Double
Private mInitialStrikeOperator                      As OptionStrikeSelectionOperators

Private mRolloverExpiryOffset                       As Long
Private mRolloverDaysBeforeExpiryToSwitch           As Long
Private mRolloverStrikeSelectionMode                As RolloverStrikeModes
Private mRolloverStrikeParameter                    As Double
Private mRolloverStrikeOperator                     As OptionStrikeSelectionOperators

Private mRolloverQuantityMode                       As RolloverQuantityModes
Private mRolloverQuantityParameter                  As BoxedDecimal
Private mRolloverQuantityLotSize                    As Long

Private mUnderlyingExchangeName                     As String

Private mCloseOrderType                             As OrderTypes
Private mCloseLimitPriceSpec                        As PriceSpecifier
Private mCloseTriggerPriceSpec                      As PriceSpecifier
Private mCloseTimeoutSecs                           As Long

Private mEntryOrderType                             As OrderTypes
Private mEntryLimitPriceSpec                        As PriceSpecifier
Private mEntryTriggerPriceSpec                      As PriceSpecifier
Private mEntryTimeoutSecs                           As Long

Private mStopLossOrderType                          As OrderTypes
Private mStopLossLimitPriceSpec                     As PriceSpecifier
Private mStopLossTriggerPriceSpec                   As PriceSpecifier

Private mTargetOrderType                            As OrderTypes
Private mTargetLimitPriceSpec                       As PriceSpecifier
Private mTargetTriggerPriceSpec                     As PriceSpecifier

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mRolloverQuantityParameter = DecimalZero
End Sub

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

Public Property Get InitialStrikeSelectionMode() As OptionStrikeSelectionModes
InitialStrikeSelectionMode = mInitialStrikeSelectionMode
End Property

Public Property Let InitialStrikeSelectionMode(ByVal Value As OptionStrikeSelectionModes)
mInitialStrikeSelectionMode = Value
End Property

Public Property Get InitialStrikeParameter() As Double
InitialStrikeParameter = mInitialStrikeParameter
End Property

Public Property Let InitialStrikeParameter(ByVal Value As Double)
mInitialStrikeParameter = Value
End Property

Public Property Get InitialStrikeOperator() As OptionStrikeSelectionOperators
InitialStrikeOperator = mInitialStrikeOperator
End Property

Public Property Let InitialStrikeOperator(ByVal Value As OptionStrikeSelectionOperators)
mInitialStrikeOperator = Value
End Property

Public Property Get RolloverDaysBeforeExpiryToSwitch() As Long
RolloverDaysBeforeExpiryToSwitch = mRolloverDaysBeforeExpiryToSwitch
End Property

Public Property Let RolloverDaysBeforeExpiryToSwitch(ByVal Value As Long)
mRolloverDaysBeforeExpiryToSwitch = Value
End Property

Public Property Get RolloverExpiryOffset() As Long
RolloverExpiryOffset = mRolloverExpiryOffset
End Property

Public Property Let RolloverExpiryOffset(ByVal Value As Long)
mRolloverExpiryOffset = Value
End Property

Public Property Get RolloverQuantityLotSize() As Long
RolloverQuantityLotSize = mRolloverQuantityLotSize
End Property

Public Property Let RolloverQuantityLotSize(ByVal Value As Long)
mRolloverQuantityLotSize = Value
End Property

Public Property Get RolloverQuantityMode() As RolloverQuantityModes
RolloverQuantityMode = mRolloverQuantityMode
End Property

Public Property Let RolloverQuantityMode(ByVal Value As RolloverQuantityModes)
mRolloverQuantityMode = Value
End Property

Public Property Get RolloverQuantityParameter() As BoxedDecimal
Set RolloverQuantityParameter = mRolloverQuantityParameter
End Property

Public Property Let RolloverQuantityParameter(ByVal Value As BoxedDecimal)
Set mRolloverQuantityParameter = Value
End Property

Public Property Get RolloverStrikeSelectionMode() As RolloverStrikeModes
RolloverStrikeSelectionMode = mRolloverStrikeSelectionMode
End Property

Public Property Let RolloverStrikeSelectionMode(ByVal Value As RolloverStrikeModes)
mRolloverStrikeSelectionMode = Value
End Property

Public Property Get RolloverStrikeOperator() As OptionStrikeSelectionOperators
RolloverStrikeOperator = mRolloverStrikeOperator
End Property

Public Property Let RolloverStrikeOperator(ByVal Value As OptionStrikeSelectionOperators)
mRolloverStrikeOperator = Value
End Property

Public Property Get RolloverStrikeParameter() As Double
RolloverStrikeParameter = mRolloverStrikeParameter
End Property

Public Property Let RolloverStrikeParameter(ByVal Value As Double)
mRolloverStrikeParameter = Value
End Property

Public Property Let StopLossOrderType(ByVal Value As OrderTypes)
mStopLossOrderType = Value
End Property

Public Property Get StopLossOrderType() As OrderTypes
StopLossOrderType = mStopLossOrderType
End Property

Public Property Get StopLossLimitPriceSpec() As PriceSpecifier
Set StopLossLimitPriceSpec = mStopLossLimitPriceSpec
End Property

Public Property Let StopLossLimitPriceSpec(ByVal Value As PriceSpecifier)
Set mStopLossLimitPriceSpec = Value
End Property

Public Property Get StopLossTriggerPriceSpec() As PriceSpecifier
Set StopLossTriggerPriceSpec = mStopLossTriggerPriceSpec
End Property

Public Property Let StopLossTriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mStopLossTriggerPriceSpec = Value
End Property

Public Property Let TargetOrderType(ByVal Value As OrderTypes)
mTargetOrderType = Value
End Property

Public Property Get TargetOrderType() As OrderTypes
TargetOrderType = mTargetOrderType
End Property

Public Property Get TargetLimitPriceSpec() As PriceSpecifier
Set TargetLimitPriceSpec = mTargetLimitPriceSpec
End Property

Public Property Let TargetLimitPriceSpec(ByVal Value As PriceSpecifier)
Set mTargetLimitPriceSpec = Value
End Property

Public Property Get TargetTriggerPriceSpec() As PriceSpecifier
Set TargetTriggerPriceSpec = mTargetTriggerPriceSpec
End Property

Public Property Let TargetTriggerPriceSpec(ByVal Value As PriceSpecifier)
Set mTargetTriggerPriceSpec = Value
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

Public Function Clone() As RolloverSpecification
Dim lRolloverSpecification As New RolloverSpecification
With Me
    Set Clone = lRolloverSpecification. _
                setDays(.Days). _
                setTime(.Time). _
                setInitialStrikeSelectionMode(.InitialStrikeSelectionMode). _
                setInitialStrikeParameter(.InitialStrikeParameter). _
                setInitialStrikeOperator(.InitialStrikeOperator). _
                setRolloverStrikeSelectionMode(.RolloverStrikeSelectionMode). _
                setRolloverStrikeParameter(.RolloverStrikeParameter). _
                setRolloverStrikeOperator(.RolloverStrikeOperator). _
                setRolloverQuantityMode(.RolloverQuantityMode). _
                setRolloverQuantityParameter(.RolloverQuantityParameter). _
                setRolloverQuantityLotSize(.RolloverQuantityLotSize). _
                setUnderlyingExchangeName(.UnderlyingExchangeName). _
                setCloseOrderType(.CloseOrderType). _
                setCloseTimeoutSecs(.CloseTimeoutSecs). _
                setCloseLimitPriceSpec(.CloseLimitPriceSpec). _
                setCloseTriggerPriceSpec(.CloseTriggerPriceSpec). _
                setEntryOrderType(.EntryOrderType). _
                setEntryTimeoutSecs(.EntryTimeoutSecs). _
                setEntryLimitPriceSpec(.EntryLimitPriceSpec). _
                setEntryTriggerPriceSpec(.EntryTriggerPriceSpec)
End With
End Function

Public Function setCloseOrderType(ByVal Value As OrderTypes) As RolloverSpecification
mCloseOrderType = Value
Set setCloseOrderType = Me
End Function

Public Function setCloseLimitPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mCloseLimitPriceSpec = Value
Set setCloseLimitPriceSpec = Me
End Function

Public Function setCloseTimeoutSecs(ByVal Value As Long) As RolloverSpecification
mCloseTimeoutSecs = Value
Set setCloseTimeoutSecs = Me
End Function

Public Function setCloseTriggerPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mCloseTriggerPriceSpec = Value
Set setCloseTriggerPriceSpec = Me
End Function

Public Function setDays(ByVal Value As Long) As RolloverSpecification
mDays = Value
Set setDays = Me
End Function

Public Function setEntryOrderType(ByVal Value As OrderTypes) As RolloverSpecification
mEntryOrderType = Value
Set setEntryOrderType = Me
End Function

Public Function setEntryLimitPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mEntryLimitPriceSpec = Value
Set setEntryLimitPriceSpec = Me
End Function

Public Function setEntryTimeoutSecs(ByVal Value As Long) As RolloverSpecification
mEntryTimeoutSecs = Value
Set setEntryTimeoutSecs = Me
End Function

Public Function setEntryTriggerPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mEntryTriggerPriceSpec = Value
Set setEntryTriggerPriceSpec = Me
End Function

Public Function setInitialStrikeSelectionMode(ByVal Value As OptionStrikeSelectionModes) As RolloverSpecification
mInitialStrikeSelectionMode = Value
Set setInitialStrikeSelectionMode = Me
End Function

Public Function setInitialStrikeParameter(ByVal Value As Double) As RolloverSpecification
mInitialStrikeParameter = Value
Set setInitialStrikeParameter = Me
End Function

Public Function setInitialStrikeOperator(ByVal Value As OptionStrikeSelectionOperators) As RolloverSpecification
mInitialStrikeOperator = Value
Set setInitialStrikeOperator = Me
End Function

Public Function setRolloverDaysBeforeExpiryToSwitch(ByVal Value As Long) As RolloverSpecification
mRolloverDaysBeforeExpiryToSwitch = Value
Set setRolloverDaysBeforeExpiryToSwitch = Me
End Function

Public Function setRolloverExpiryOffset(ByVal Value As Long) As RolloverSpecification
mRolloverExpiryOffset = Value
Set setRolloverExpiryOffset = Me
End Function

Public Function setRolloverQuantityMode(ByVal Value As RolloverQuantityModes) As RolloverSpecification
mRolloverQuantityMode = Value
Set setRolloverQuantityMode = Me
End Function

Public Function setRolloverQuantityParameter(ByVal Value As BoxedDecimal) As RolloverSpecification
Set mRolloverQuantityParameter = Value
Set setRolloverQuantityParameter = Me
End Function

Public Function setRolloverQuantityLotSize(ByVal Value As Long) As RolloverSpecification
mRolloverQuantityLotSize = Value
Set setRolloverQuantityLotSize = Me
End Function

Public Function setRolloverStrikeSelectionMode(ByVal Value As RolloverStrikeModes) As RolloverSpecification
mRolloverStrikeSelectionMode = Value
Set setRolloverStrikeSelectionMode = Me
End Function

Public Function setRolloverStrikeOperator(ByVal Value As OptionStrikeSelectionOperators) As RolloverSpecification
mRolloverStrikeOperator = Value
Set setRolloverStrikeOperator = Me
End Function

Public Function setRolloverStrikeParameter(ByVal Value As Double) As RolloverSpecification
mRolloverStrikeParameter = Value
Set setRolloverStrikeParameter = Me
End Function

Public Function setStopLossOrderType(ByVal Value As OrderTypes) As RolloverSpecification
mStopLossOrderType = Value
Set setStopLossOrderType = Me
End Function

Public Function setStopLossLimitPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mStopLossLimitPriceSpec = Value
Set setStopLossLimitPriceSpec = Me
End Function

Public Function setStopLossTriggerPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mStopLossTriggerPriceSpec = Value
Set setStopLossTriggerPriceSpec = Me
End Function

Public Function setTargetOrderType(ByVal Value As OrderTypes) As RolloverSpecification
mTargetOrderType = Value
Set setTargetOrderType = Me
End Function

Public Function setTargetLimitPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mTargetLimitPriceSpec = Value
Set setTargetLimitPriceSpec = Me
End Function

Public Function setTargetTriggerPriceSpec(ByVal Value As PriceSpecifier) As RolloverSpecification
Set mTargetTriggerPriceSpec = Value
Set setTargetTriggerPriceSpec = Me
End Function

Public Function setTime(ByVal Value As Date) As RolloverSpecification
mTime = Value
Set setTime = Me
End Function

Public Function setUnderlyingExchangeName(ByVal Value As String) As RolloverSpecification
mUnderlyingExchangeName = Value
Set setUnderlyingExchangeName = Me
End Function

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "Days", mDays, 0

jb.EncodeNameAndValue "Time", CDbl(mTime), 0

jb.EncodeNameAndValue "InitialStrikeSelectionMode", mInitialStrikeSelectionMode, OptionStrikeSelectionModes.OptionStrikeSelectionModeNone
jb.EncodeNameAndValue "InitialStrikeParameter", mInitialStrikeParameter, 0
jb.EncodeNameAndValue "InitialStrikeOperator", mInitialStrikeOperator, OptionStrikeSelectionOperatorNone

jb.EncodeNameAndValue "RolloverDaysBeforeExpiryToSwitch", mRolloverDaysBeforeExpiryToSwitch, 1
jb.EncodeNameAndValue "RolloverExpiryOffset", mRolloverExpiryOffset, 1
jb.EncodeNameAndValue "RolloverStrikeSelectionMode", mRolloverStrikeSelectionMode, RolloverStrikeModeNone
jb.EncodeNameAndValue "RolloverStrikeParameter", mRolloverStrikeParameter, 0
jb.EncodeNameAndValue "RolloverStrikeOperator", mRolloverStrikeOperator, OptionStrikeSelectionOperatorNone

jb.EncodeNameAndValue "RolloverQuantityMode", mRolloverQuantityMode, RolloverQuantityModes.RolloverQuantityModeNone
jb.EncodeNameAndValue "RolloverQuantityParameter", mRolloverQuantityParameter, DecimalZero
jb.EncodeNameAndValue "RolloverQuantityLotSize", mRolloverQuantityLotSize, 0

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

jb.EncodeNameAndValue "StopLossOrderType", mStopLossOrderType, OrderTypeNone
If Not gIsNullPriceSpecifier(mStopLossLimitPriceSpec) Then _
    jb.EncodeNameAndValue "StopLossLimitPriceSpec", mStopLossLimitPriceSpec, Nothing
If Not gIsNullPriceSpecifier(mStopLossTriggerPriceSpec) Then _
    jb.EncodeNameAndValue "StopLossTriggerPriceSpec", mStopLossTriggerPriceSpec, Nothing

jb.EncodeNameAndValue "TargetOrderType", mTargetOrderType, OrderTypeNone
If Not gIsNullPriceSpecifier(mTargetLimitPriceSpec) Then _
    jb.EncodeNameAndValue "TargetLimitPriceSpec", mTargetLimitPriceSpec, Nothing
If Not gIsNullPriceSpecifier(mTargetTriggerPriceSpec) Then _
    jb.EncodeNameAndValue "TargetTriggerPriceSpec", mTargetTriggerPriceSpec, Nothing

jb.EndObject

ToJSON = jb.JSONString
End Function

'@================================================================================
' Helper Functions
'@================================================================================




