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

Private mRs As ADODB.Recordset
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

Friend Property Get Bookmark() As Variant
Bookmark = mBookmark
End Property

Public Property Get LocalSymbol() As String
Const ProcName As String = "LocalSymbol"

On Error GoTo Err

LocalSymbol = Trim$(mRs("LocalSymbol"))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LocalSymbol( _
                ByVal value As String)
Const ProcName As String = "LocalSymbol"

On Error GoTo Err

value = Trim$(value)
mRs("LocalSymbol") = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ProviderKey() As String
Const ProcName As String = "ProviderKey"

On Error GoTo Err

ProviderKey = Trim$(mRs("providerkey"))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ProviderKey( _
                ByVal value As String)
Const ProcName As String = "ProviderKey"

On Error GoTo Err

value = Trim$(value)
mRs("providerkey") = value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub Initialise( _
                ByVal value As ADODB.Recordset, _
                ByVal Bookmark As Variant)
Set mRs = value
mBookmark = Bookmark
mRs.Bookmark = mBookmark
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



