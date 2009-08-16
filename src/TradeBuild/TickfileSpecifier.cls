VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements ITickfileSpecifier

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

Private Const ProjectName                   As String = "TradeBuild26"
Private Const ModuleName                    As String = "TickfileSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileFormatID       As String
Private mFilename               As String
Private mContract               As Contract
Private mFromDate               As Date
Private mToDate                 As Date
Private mEntireSession          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITickfileSpecifier Interface Members
'@================================================================================

Private Property Let ITickfileSpecifier_Contract(ByVal RHS As Contract)
Set mContract = RHS
End Property

Private Property Get ITickfileSpecifier_Contract() As Contract
Set ITickfileSpecifier_Contract = mContract
End Property

Private Property Let ITickfileSpecifier_EntireSession(ByVal RHS As Boolean)
mEntireSession = RHS
End Property

Private Property Get ITickfileSpecifier_EntireSession() As Boolean
ITickfileSpecifier_EntireSession = mEntireSession
End Property

Private Property Let ITickfileSpecifier_Filename(ByVal RHS As String)
mFilename = RHS
End Property

Private Property Get ITickfileSpecifier_Filename() As String
ITickfileSpecifier_Filename = mFilename
End Property

Private Property Let ITickfileSpecifier_FromDate(ByVal RHS As Date)
mFromDate = RHS
End Property

Private Property Get ITickfileSpecifier_FromDate() As Date
ITickfileSpecifier_FromDate = mFromDate
End Property

Private Property Let ITickfileSpecifier_TickfileFormatID(ByVal RHS As String)
mTickfileFormatID = RHS
End Property

Private Property Get ITickfileSpecifier_TickfileFormatID() As String
ITickfileSpecifier_TickfileFormatID = mTickfileFormatID
End Property

Private Property Let ITickfileSpecifier_ToDate(ByVal RHS As Date)
mToDate = RHS
End Property

Private Property Get ITickfileSpecifier_ToDate() As Date
ITickfileSpecifier_ToDate = mToDate
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Contract(ByVal value As Contract)
Set mContract = value
End Property

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Public Property Let EntireSession(ByVal value As Boolean)
mEntireSession = value
End Property

Public Property Get EntireSession() As Boolean
EntireSession = mEntireSession
End Property

Public Property Let Filename(ByVal value As String)
mFilename = value
End Property

Public Property Get Filename() As String
Filename = mFilename
End Property

Public Property Let FromDate(ByVal value As Date)
mFromDate = value
End Property

Public Property Get FromDate() As Date
FromDate = mFromDate
End Property

Public Property Let TickfileFormatID(ByVal value As String)
mTickfileFormatID = value
End Property

Public Property Get TickfileFormatID() As String
TickfileFormatID = mTickfileFormatID
End Property

Public Property Let ToDate(ByVal value As Date)
mToDate = value
End Property

Public Property Get ToDate() As Date
ToDate = mToDate
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


