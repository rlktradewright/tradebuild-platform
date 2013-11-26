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


Private Const ModuleName                    As String = "TickfileSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileFormatID       As String
Private mFilename               As String
Private mContractFuture         As IFuture
Private mFromDate               As Date
Private mToDate                 As Date
Private mEntireSession          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITickfileSpecifier Interface Members
'@================================================================================

Private Property Let ITickfileSpecifier_ContractFuture(ByVal RHS As IFuture)
Set mContractFuture = RHS
End Property

Private Property Get ITickfileSpecifier_ContractFuture() As IFuture
Set ITickfileSpecifier_ContractFuture = mContractFuture
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

Public Property Let ContractFuture(ByVal Value As IFuture)
Set mContractFuture = Value
End Property

Public Property Get ContractFuture() As IFuture
Set ContractFuture = mContractFuture
End Property

Public Property Let EntireSession(ByVal Value As Boolean)
mEntireSession = Value
End Property

Public Property Get EntireSession() As Boolean
EntireSession = mEntireSession
End Property

Public Property Let Filename(ByVal Value As String)
mFilename = Value
End Property

Public Property Get Filename() As String
Filename = mFilename
End Property

Public Property Let FromDate(ByVal Value As Date)
mFromDate = Value
End Property

Public Property Get FromDate() As Date
FromDate = mFromDate
End Property

Public Property Let TickfileFormatID(ByVal Value As String)
mTickfileFormatID = Value
End Property

Public Property Get TickfileFormatID() As String
TickfileFormatID = mTickfileFormatID
End Property

Public Property Let ToDate(ByVal Value As Date)
mToDate = Value
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


