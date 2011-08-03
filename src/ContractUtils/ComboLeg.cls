VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboLeg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "ComboLeg"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec As ContractSpecifier
Private mIsBuyLeg As Boolean
Private mRatio As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ContractSpec() As ContractSpecifier
Set ContractSpec = mContractSpec
End Property

Public Property Get IsBuyLeg() As Boolean
IsBuyLeg = mIsBuyLeg
End Property

Public Property Get Ratio() As Long
Ratio = mRatio
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractSpec As ContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long)
Set mContractSpec = pContractSpec
mIsBuyLeg = pIsBuyLeg
mRatio = pRatio
End Sub

Public Function ToString() As String

Const ProcName As String = "ToString"
Dim failpoint As String
On Error GoTo Err

ToString = "action=" & IIf(mIsBuyLeg, "BUY", "SELL") & "; " & _
            "ratio=" & mRatio & "; " & _
            "contractSpec=(" & mContractSpec.ToString & "); "

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint

End Function

'@================================================================================
' Helper Functions
'@================================================================================


