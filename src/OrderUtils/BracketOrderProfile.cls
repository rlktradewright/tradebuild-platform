VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BracketOrderProfile"
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

Private Const ModuleName                            As String = "BracketOrderProfile"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec                               As IContractSpecifier
Private mKey                                        As String
Private mAction                                     As OrderActions
Private mQuantity                                   As Long
Private mEntryPrice                                 As Double
Private mExitPrice                                  As Double
Private mProfit                                     As Currency
Private mMaxProfit                                  As Currency
Private mMaxLoss                                    As Currency
Private mRisk                                       As Currency
Private mStartTime                                  As Date
Private mEndTime                                    As Date
Private mDescription                                As String
Private mEntryReason                                As String
Private mStopReason                                 As String
Private mTargetReason                               As String
Private mTimezoneName                               As String

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

Public Property Get Action() As OrderActions
Action = mAction
End Property

Public Property Get Description() As String
Description = mDescription
End Property

Public Property Get EndTime() As Date
EndTime = mEndTime
End Property

Public Property Get EntryPrice() As Double
EntryPrice = mEntryPrice
End Property

Public Property Get EntryReason() As String
EntryReason = mEntryReason
End Property

Public Property Get ExitPrice() As Double
ExitPrice = mExitPrice
End Property

Public Property Get Key() As String
Key = mKey
End Property

Public Property Get MaxLoss() As Currency
MaxLoss = mMaxLoss
End Property

Public Property Get MaxProfit() As Currency
MaxProfit = mMaxProfit
End Property

Public Property Get Profit() As Currency
Profit = mProfit
End Property

Public Property Get Quantity() As Long
Quantity = mQuantity
End Property

Public Property Get Risk() As Currency
Risk = mRisk
End Property

Public Property Get StartTime() As Date
StartTime = mStartTime
End Property

Public Property Get StopReason() As String
StopReason = mStopReason
End Property

Public Property Get TargetReason() As String
TargetReason = mTargetReason
End Property

Public Property Get TimezoneName() As String
TimezoneName = mTimezoneName
End Property
Friend Property Let Key(ByVal Value As String)
mKey = Value
End Property

Friend Property Let Action(ByVal Value As OrderActions)
mAction = Value
End Property

Friend Property Let ContractSpec(ByVal Value As IContractSpecifier)
Set mContractSpec = Value
End Property

Friend Property Let Description(ByVal Value As String)
mDescription = Value
End Property

Friend Property Let EndTime(ByVal Value As Date)
mEndTime = Value
End Property

Friend Property Let EntryPrice(ByVal Value As Double)
mEntryPrice = Value
End Property

Friend Property Let EntryReason(ByVal Value As String)
mEntryReason = Value
End Property

Friend Property Let ExitPrice(ByVal Value As Double)
mExitPrice = Value
End Property

Friend Property Let MaxLoss(ByVal Value As Currency)
mMaxLoss = Value
End Property

Friend Property Let MaxProfit(ByVal Value As Currency)
mMaxProfit = Value
End Property

Friend Property Let Profit(ByVal Value As Currency)
mProfit = Value
End Property

Friend Property Let Quantity(ByVal Value As Long)
mQuantity = Value
End Property

Friend Property Let Risk(ByVal Value As Currency)
mRisk = Value
End Property

Friend Property Let StartTime(ByVal Value As Date)
mStartTime = Value
End Property

Friend Property Let StopReason(ByVal Value As String)
mStopReason = Value
End Property

Friend Property Let TargetReason(ByVal Value As String)
mTargetReason = Value
End Property

Friend Property Let TimezoneName(ByVal Value As String)
mTimezoneName = Value
End Property


'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================






