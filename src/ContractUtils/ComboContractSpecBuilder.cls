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

Private Const ProjectName                   As String = "ContractUtils26"
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
mContractSpecifier.sectype = SecTypeCombo

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

Public Function addLeg( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal isBuyLeg As Boolean, _
                ByVal ratio As Long, _
                Optional ByVal key As String) As ComboLeg
Set addLeg = mComboLegs.addLeg(contractSpec, isBuyLeg, ratio, key)
End Function

Public Sub removeLeg(ByVal index As Variant)
mComboLegs.remove index
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


