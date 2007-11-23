VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
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

Private Const ProjectName                   As String = "gcd"
Private Const ModuleName                    As String = "ContractProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec As ContractSpecifier
Private WithEvents mContracts As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mDone As Boolean
Private mLineNumber As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
gCon.writeErrorLine "Line " & mLineNumber & ": Invalid contract spec " & mContractSpec.ToString
End Sub

Private Sub mContracts_NoMoreContractDetails()
mDone = True
End Sub

Private Sub mContracts_Notification(ev As TWUtilities30.NotificationEvent)
gCon.writeErrorLine "Line " & mLineNumber & ": Notification: " & ev.eventMessage
End Sub

Private Sub mContracts_TryLater(ByVal reason As String, ByVal serviceProviderName As String)
gCon.writeErrorLine "Line " & mLineNumber & ": Try later: " & reason
End Sub
'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub process( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal lineNumber As Long)
Dim lContract As Contract

mDone = False

Set mContractSpec = contractSpec
mLineNumber = lineNumber
Set mContracts = TradeBuildAPI.loadContracts(contractSpec)

' now hang around until we've got the info, but don't exit this proc
' because we don't want to process the next line of input till this one's done

Do While Not mDone
    Wait (100)
Loop

If mContracts.Count <> 0 Then
    For Each lContract In mContracts
        gCon.writeLine SecTypeToString(lContract.specifier.sectype) & "," & _
                        lContract.specifier.exchange & "," & _
                        lContract.specifier.localSymbol & "," & _
                        lContract.specifier.symbol & "," & _
                        lContract.specifier.currencyCode & "," & _
                        lContract.specifier.expiry & "," & _
                        lContract.specifier.strike & "," & _
                        OptionRightToString(lContract.specifier.right)
    Next
End If

End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================


