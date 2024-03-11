VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LiveOrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "LiveOrderEntry"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestamp                                  As Date
Private mKey                                        As String
Private mOrder                                      As IBracketOrder
Private mBracketOrderSpec                           As BracketOrderSpecification
Private mGroupName                                  As String
Private mCancelled                                  As Boolean

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

Friend Property Let BracketOrderSpec(ByVal Value As BracketOrderSpecification)
Set mBracketOrderSpec = Value
End Property

Friend Property Get BracketOrderSpec() As BracketOrderSpecification
Set BracketOrderSpec = mBracketOrderSpec
End Property

Friend Property Let Cancelled(ByVal Value As Boolean)
mCancelled = Value
End Property

Friend Property Get Cancelled() As Boolean
Cancelled = mCancelled
End Property

Friend Property Let GroupName(ByVal Value As String)
mGroupName = Value
End Property

Friend Property Get GroupName() As String
GroupName = mGroupName
End Property

Friend Property Let Key(ByVal Value As String)
mKey = Value
End Property

Friend Property Get Key() As String
Key = mKey
End Property

Friend Property Let Order(ByVal Value As IBracketOrder)
Set mOrder = Value
End Property

Friend Property Get Order() As IBracketOrder
Set Order = mOrder
End Property

Friend Property Let Timestamp(ByVal Value As Date)
mTimestamp = Value
End Property

Friend Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




