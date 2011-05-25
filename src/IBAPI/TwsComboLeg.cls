VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TwsComboLeg"
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

Private Const ModuleName                            As String = "TwsComboLeg"

'@================================================================================
' Member variables
'@================================================================================

Private mConId                                        As Long
Private mRatio                                        As Long
Private mAction                                       As TwsOrderActions
Private mExchange                                     As String
Private mOpenClose                                    As TwsLegOpenCloseCodes

' for stock legs when doing short sale
Private mShortSaleSlot                                As TwsShortSaleSlotCodes
Private mDesignatedLocation                           As String

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

Public Property Let Action(ByVal Value As TwsOrderActions)
mAction = Value
End Property

Public Property Get Action() As TwsOrderActions
Action = mAction
End Property

Public Property Let ConId(ByVal Value As Long)
mConId = Value
End Property

Public Property Get ConId() As Long
ConId = mConId
End Property

Public Property Let DesignatedLocation(ByVal Value As String)
mDesignatedLocation = Value
End Property

Public Property Get DesignatedLocation() As String
DesignatedLocation = mDesignatedLocation
End Property

Public Property Let Exchange(ByVal Value As String)
mExchange = Value
End Property

Public Property Get Exchange() As String
Exchange = mExchange
End Property

Public Property Let OpenClose(ByVal Value As TwsLegOpenCloseCodes)
mOpenClose = Value
End Property

Public Property Get OpenClose() As TwsLegOpenCloseCodes
OpenClose = mOpenClose
End Property

Public Property Let Ratio(ByVal Value As Long)
mRatio = Value
End Property

Public Property Get Ratio() As Long
Ratio = mRatio
End Property

Public Property Let ShortSaleSlot(ByVal Value As TwsShortSaleSlotCodes)
mShortSaleSlot = Value
End Property

Public Property Get ShortSaleSlot() As TwsShortSaleSlotCodes
ShortSaleSlot = mShortSaleSlot
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




