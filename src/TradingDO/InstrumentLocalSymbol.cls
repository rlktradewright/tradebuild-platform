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

Private mRS As ADODB.Recordset
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

LocalSymbol = Trim$(mRS("LocalSymbol"))

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LocalSymbol( _
                ByVal Value As String)
Const ProcName As String = "LocalSymbol"

On Error GoTo Err

Value = Trim$(Value)
mRS("LocalSymbol") = Value

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ProviderKey() As String
Const ProcName As String = "ProviderKey"

On Error GoTo Err

ProviderKey = Trim$(mRS("providerkey"))

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ProviderKey( _
                ByVal Value As String)
Const ProcName As String = "ProviderKey"

On Error GoTo Err

Value = Trim$(Value)
mRS("providerkey") = Value

Exit Property

Err:
GTDO.HandleUnexpectedError ProcName, ModuleName
End Property

'@===============================================================================
' Methods
'@===============================================================================

Friend Sub Initialise( _
                ByVal Value As ADODB.Recordset, _
                ByVal Bookmark As Variant)
Set mRS = Value
mBookmark = Bookmark
mRS.Bookmark = mBookmark
End Sub

'@===============================================================================
' Helper Functions
'@===============================================================================



