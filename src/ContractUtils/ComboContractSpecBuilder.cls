VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboContractSpecBldr"
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

Private Const ModuleName                    As String = "ComboContractSpecBldr"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpecifier                  As ContractSpecifier
Private mComboLegs                          As ComboLegs

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mContractSpecifier = New ContractSpecifier
mContractSpecifier.zSecType = SecTypeCombo

Set mComboLegs = New ComboLegs
mContractSpecifier.ComboLegs = mComboLegs
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractSpecifier() As IContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddLeg( _
                ByVal ContractSpec As IContractSpecifier, _
                ByVal IsBuyLeg As Boolean, _
                ByVal Ratio As Long) As ComboLeg
Const ProcName As String = "AddLeg"

On Error GoTo Err

Set AddLeg = mComboLegs.AddLeg(ContractSpec, IsBuyLeg, Ratio)

Exit Function

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveLeg(ByVal index As Variant)
Const ProcName As String = "RemoveLeg"

On Error GoTo Err

mComboLegs.Remove index

Exit Sub

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


