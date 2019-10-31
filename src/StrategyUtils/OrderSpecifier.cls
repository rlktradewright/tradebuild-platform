VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSpecifier"
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

Private Const ModuleName                            As String = "OrderSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mOrderRole                                  As OrderRoles
Private mOrderType                                  As Long
Private mPriceSpec                                  As PriceSpecifier
Private mTriggerPriceSpec                           As PriceSpecifier
Private mTIF                                        As Long
Private mIgnoreRTH                                  As Boolean
Private mIsSimulated                                As Boolean

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

Friend Property Get IgnoreRTH() As Boolean
IgnoreRTH = mIgnoreRTH
End Property

Friend Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Get OrderRole() As OrderRoles
OrderRole = mOrderRole
End Property

Friend Property Get OrderType() As Long
OrderType = mOrderType
End Property

Friend Property Get PriceSpecifier() As PriceSpecifier
Set PriceSpecifier = mPriceSpec
End Property

Friend Property Get TriggerPriceSpecifier() As PriceSpecifier
Set TriggerPriceSpecifier = mTriggerPriceSpec
End Property

Friend Property Get TIF() As Long
TIF = mTIF
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pOrderRole As OrderRoles, _
                ByVal pOrderType As Long, _
                ByVal pIsSimulated As Boolean, _
                ByVal pPriceSpec As PriceSpecifier, _
                ByVal pTriggerPriceSpec As PriceSpecifier, _
                ByVal pTIF As Long, _
                ByVal pIgnoreRTH As Boolean)
mOrderRole = pOrderRole
mOrderType = pOrderType
mIsSimulated = pIsSimulated
Set mPriceSpec = pPriceSpec
Set mTriggerPriceSpec = pTriggerPriceSpec
mTIF = pTIF
mIgnoreRTH = pIgnoreRTH
End Sub


'@================================================================================
' Helper Functions
'@================================================================================




