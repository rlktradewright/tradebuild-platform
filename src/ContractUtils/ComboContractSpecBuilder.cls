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
mContractSpecifier.SecType = SecTypeCombo

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

Public Property Get ContractSpecifier() As ContractSpecifier
Set ContractSpecifier = mContractSpecifier
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function AddLeg( _
                ByVal ContractSpec As ContractSpecifier, _
                ByVal IsBuyLeg As Boolean, _
                ByVal ratio As Long, _
                Optional ByVal Key As String) As ComboLeg
Const ProcName As String = "AddLeg"
Dim failpoint As String
On Error GoTo Err

Set AddLeg = mComboLegs.AddLeg(ContractSpec, IsBuyLeg, ratio, Key)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Sub RemoveLeg(ByVal index As Variant)
Const ProcName As String = "RemoveLeg"
Dim failpoint As String
On Error GoTo Err

mComboLegs.Remove index

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


