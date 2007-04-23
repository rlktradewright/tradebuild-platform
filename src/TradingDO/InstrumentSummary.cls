VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "InstrumentSummary"

'@================================================================================
' Member variables
'@================================================================================

Private mIdentity As Long
Private mName As String
Private mSymbol As String
Private mExpiryMonth As String
Private mCategory As String

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

Public Property Get id() As Long
id = mIdentity
End Property

Friend Property Let id(ByVal value As Long)
mIdentity = value
End Property

Public Property Get name() As String
name = mName
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get symbol() As String
symbol = mSymbol
End Property

Friend Property Let symbol(ByVal value As String)
mSymbol = value
End Property

Public Property Get expiryMonth() As String
expiryMonth = mExpiryMonth
End Property

Friend Property Let expiryMonth(ByVal value As String)
mExpiryMonth = value
End Property

Public Property Get category() As String
category = mCategory
End Property

Friend Property Let category(ByVal value As String)
mCategory = value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================






