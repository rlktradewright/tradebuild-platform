VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsOrderState"
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

Private Const ModuleName                            As String = "TwsOrderState"

'@================================================================================
' Member variables
'@================================================================================

Private mStatus                                     As String

Private mInitMargin                                 As String
Private mMaintMargin                                As String
Private mEquityWithLoan                             As String

Private mCommission                                 As Double
Private mMinCommission                              As Double
Private mMaxCommission                              As Double
Private mCommissionCurrency                         As String
    
Private mWarningText                                As String

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

Public Property Let Status(ByVal Value As String)
mStatus = Value
End Property

Public Property Get Status() As String
Status = mStatus
End Property

Public Property Let InitMargin(ByVal Value As String)
mInitMargin = Value
End Property

Public Property Get InitMargin() As String
InitMargin = mInitMargin
End Property

Public Property Let MaintMargin(ByVal Value As String)
mMaintMargin = Value
End Property

Public Property Get MaintMargin() As String
MaintMargin = mMaintMargin
End Property

Public Property Let EquityWithLoan(ByVal Value As String)
mEquityWithLoan = Value
End Property

Public Property Get EquityWithLoan() As String
EquityWithLoan = mEquityWithLoan
End Property

Public Property Let Commission(ByVal Value As Double)
mCommission = Value
End Property

Public Property Get Commission() As Double
Commission = mCommission
End Property

Public Property Let MinCommission(ByVal Value As Double)
mMinCommission = Value
End Property

Public Property Get MinCommission() As Double
MinCommission = mMinCommission
End Property

Public Property Let MaxCommission(ByVal Value As Double)
mMaxCommission = Value
End Property

Public Property Get MaxCommission() As Double
MaxCommission = mMaxCommission
End Property

Public Property Let CommissionCurrency(ByVal Value As String)
mCommissionCurrency = Value = Value
End Property

Public Property Get CommissionCurrency() As String
CommissionCurrency = mCommissionCurrency
End Property

Public Property Let WarningText(ByVal Value As String)
mWarningText = Value
End Property

Public Property Get WarningText() As String
WarningText = mWarningText
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




