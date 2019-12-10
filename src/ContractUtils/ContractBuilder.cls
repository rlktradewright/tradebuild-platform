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

Public Property Let DaysBeforeExpiryToSwitch(ByVal Value As Long)
mContract.DaysBeforeExpiryToSwitch = Value
End Property

Public Property Let Description(ByVal Value As String)
mContract.Description = Value
End Property

Public Property Let ExpiryDate(ByVal Value As Date)
mContract.ExpiryDate = Value
End Property

Public Property Let ProviderIDs(ByVal Value As Parameters)
mContract.ProviderIDs = Value
End Property

Public Property Let ProviderProperties(ByVal Value As Parameters)
mContract.ProviderProperties = Value
End Property

Public Property Let SessionStartTime(ByVal Value As Date)
mContract.SessionStartTime = TimeValue(Value)
End Property

Public Property Let SessionEndTime(ByVal Value As Date)
mContract.SessionEndTime = TimeValue(Value)
End Property

Public Property Let TickSize(ByVal Value As Double)
mContract.TickSize = Value
End Property

Public Property Let TimezoneName( _
                ByVal Value As String)
mContract.TimezoneName = Value
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


