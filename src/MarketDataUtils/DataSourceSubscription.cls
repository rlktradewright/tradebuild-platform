VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceSubscription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "DataSourceSubscription"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                 As MarketDataSource
Private mSubscribedToMarketData                     As Boolean
Private mSubscribedToMarketDepth                    As Boolean

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

Friend Property Let DataSource(ByVal Value As MarketDataSource)
Set mDataSource = Value
End Property

Friend Property Get DataSource() As MarketDataSource
Set DataSource = mDataSource
End Property

Friend Property Let SubscribedToMarketData(ByVal Value As Boolean)
mSubscribedToMarketData = Value
End Property

Friend Property Get SubscribedToMarketData() As Boolean
SubscribedToMarketData = mSubscribedToMarketData
End Property

Friend Property Let SubscribedToMarketDepth(ByVal Value As Boolean)
mSubscribedToMarketDepth = Value
End Property

Friend Property Get SubscribedToMarketDepth() As Boolean
SubscribedToMarketDepth = mSubscribedToMarketDepth
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================




