VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractBuilder"
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

Private Const ProjectName                   As String = "ContractUtils26"
Private Const ModuleName                    As String = "ContractBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mContract                           As Contract

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

Public Property Get Contract() As Contract
Set Contract = mContract
End Property

Friend Property Let daysBeforeExpiryToSwitch(ByVal value As Long)
mContract.daysBeforeExpiryToSwitch = value
End Property

Friend Property Let description(ByVal value As String)
mContract.description = value
End Property

Friend Property Let expiryDate(ByVal value As Date)
mContract.expiryDate = value
End Property

Friend Property Let multiplier(ByVal value As Long)
mContract.multiplier = value
End Property

Friend Property Let providerIDs(ByRef value As Parameters)
mContract.providerIDs = value
End Property

Friend Property Let sessionStartTime(ByVal value As Date)
mContract.sessionStartTime = TimeValue(value)
End Property

Friend Property Let sessionEndTime(ByVal value As Date)
mContract.sessionEndTime = TimeValue(value)
End Property

Friend Property Let TickSize(ByVal value As Double)
mContract.TickSize = value
End Property

Friend Property Let timezone( _
                ByVal value As timezone)
Set mContract.timezone = value
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal specifier As ContractSpecifier)
Set mContract = New Contract
mContract.specifier = specifier
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


