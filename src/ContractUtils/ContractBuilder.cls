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

Public Property Get Contract() As IContract
Set Contract = mContract
End Property

Public Property Let DaysBeforeExpiryToSwitch(ByVal value As Long)
mContract.DaysBeforeExpiryToSwitch = value
End Property

Public Property Let Description(ByVal value As String)
mContract.Description = value
End Property

Public Property Let ExpiryDate(ByVal value As Date)
mContract.ExpiryDate = value
End Property

Public Property Let Multiplier(ByVal value As Double)
mContract.Multiplier = value
End Property

Public Property Let ProviderIDs(ByVal value As Parameters)
mContract.ProviderIDs = value
End Property

Public Property Let SessionStartTime(ByVal value As Date)
mContract.SessionStartTime = TimeValue(value)
End Property

Public Property Let SessionEndTime(ByVal value As Date)
mContract.SessionEndTime = TimeValue(value)
End Property

Public Property Let TickSize(ByVal value As Double)
mContract.TickSize = value
End Property

Public Property Let TimezoneName( _
                ByVal value As String)
mContract.TimezoneName = value
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub BuildFrom( _
                ByVal pContract As IContract)
Const ProcName As String = "BuildFrom"

On Error GoTo Err

With mContract
    .DaysBeforeExpiryToSwitch = pContract.DaysBeforeExpiryToSwitch
    .Description = pContract.Description
    .ExpiryDate = pContract.ExpiryDate
    .Multiplier = pContract.Multiplier
    '.ProviderIDs = pContract.ProviderIDs
    .SessionEndTime = pContract.SessionEndTime
    .SessionStartTime = pContract.SessionStartTime
    .Specifier = pContract.Specifier
    .TickSize = pContract.TickSize
    .TimezoneName = pContract.TimezoneName
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub Initialise( _
                ByVal Specifier As IContractSpecifier)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mContract = New Contract
mContract.Specifier = Specifier
mContract.TimezoneName = ""
mContract.DaysBeforeExpiryToSwitch = 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
                
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mContract = gLoadContractFromConfig(pConfig)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


