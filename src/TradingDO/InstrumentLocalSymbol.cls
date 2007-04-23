VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InstrumentLocalSymbol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TradingDO26"
Private Const ModuleName As String = "InstrumentLocalSymbol"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mRs As ADODB.recordset
Private mBookmark As Variant

'@===============================================================================
' Class Event Handlers
'@===============================================================================

'@===============================================================================
' XXXX Interface Members
'@===============================================================================

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@===============================================================================
' Properties
'@===============================================================================

Friend Property Get bookmark() As Variant
bookmark = mBookmark
End Property

Public Property Get localSymbol() As String
localSymbol = Trim$(mRs("localSymbol"))
End Property

Public Property Let localSymbol( _
                ByVal value As String)
mRs("localSymbol") = value
End Property

Public Property Get providerKey() As String
providerKey = Trim$(mRs("providerkey"))
End Property

Public Property Let providerKey( _
                ByVal value As String)
mRs("providerkey") = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub initialise( _
                ByVal value As ADODB.recordset, _
                ByVal bookmark As Variant)
Set mRs = value
mBookmark = bookmark
mRs.bookmark = mBookmark
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



