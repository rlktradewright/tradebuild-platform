VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RolloverSpecification"
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

Implements IJSONable

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

Private Const ModuleName                            As String = "RolloverSpecification"

'@================================================================================
' Member variables
'@================================================================================

Private mDays                                       As Long
Private mTime                                       As Date
Private mCloseAtMarket                              As Boolean
Private mCloseSpreadfactor                          As Long
Private mEntryAtMarket                              As Boolean
Private mEntrySpreadfactor                          As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Days(ByVal Value As Long)
mDays = Value
End Property

Public Property Get Days() As Long
Days = mDays
End Property

Public Property Let Time(ByVal Value As Date)
mTime = Value
End Property

Public Property Get Time() As Date
Time = mTime
End Property

Public Property Let CloseAtMarket(ByVal Value As Boolean)
mCloseAtMarket = Value
End Property

Public Property Get CloseAtMarket() As Boolean
CloseAtMarket = mCloseAtMarket
End Property

Public Property Let CloseSpreadfactor(ByVal Value As Long)
mCloseSpreadfactor = Value
End Property

Public Property Get CloseSpreadfactor() As Long
CloseSpreadfactor = mCloseSpreadfactor
End Property

Public Property Let EntryAtMarket(ByVal Value As Boolean)
mEntryAtMarket = Value
End Property

Public Property Get EntryAtMarket() As Boolean
EntryAtMarket = mEntryAtMarket
End Property

Public Property Let EntrySpreadfactor(ByVal Value As Long)
mEntrySpreadfactor = Value
End Property

Public Property Get EntrySpreadfactor() As Long
EntrySpreadfactor = mEntrySpreadfactor
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pDays As Long, _
                ByVal pTime As Date, _
                ByVal pCloseAtMarket As Boolean, _
                ByVal pCloseSpreadfactor As Long, _
                ByVal pEntryAtMarket As Boolean, _
                ByVal pEntrySpreadfactor As Long)
mDays = pDays
mTime = pTime
mCloseAtMarket = pCloseAtMarket
mCloseSpreadfactor = pCloseSpreadfactor
mEntryAtMarket = pEntryAtMarket
mEntrySpreadfactor = pEntrySpreadfactor
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "CloseAtMarket", mCloseAtMarket, False
jb.EncodeNameAndValue "CloseSpreadfactor", mCloseSpreadfactor, 0

jb.EncodeNameAndValue "Days", mDays, 0

jb.EncodeNameAndValue "EntryAtMarket", mEntryAtMarket, False
jb.EncodeNameAndValue "EntrySpreadfactor", mEntrySpreadfactor, 0

jb.EncodeNameAndValue "Time", CDbl(mTime), 0

jb.EndObject

ToJSON = jb.JSONString
End Function




