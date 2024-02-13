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

Implements IComboleg
Implements IJSONable
Implements IStringable

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
' IComboleg Interface Members
'@================================================================================

Private Property Get IComboleg_ContractSpec() As IContractSpecifier
Set IComboleg_ContractSpec = ContractSpec
End Property

Private Property Get IComboleg_IsBuyLeg() As Boolean
IComboleg_IsBuyLeg = IsBuyLeg
End Property

Private Property Get IComboleg_Ratio() As Long
IComboleg_Ratio = Ratio
End Property

Private Function IComboleg_ToString() As String
IComboleg_ToString = ToString
End Function

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJsonable_ToJSON() As String
IJsonable_ToJSON = ToJSON
End Function

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let zContractSpec(ByVal Value As IContractSpecifier)
Attribute zContractSpec.VB_MemberFlags = "40"
Set mContractSpec = Value
End Property

Public Property Get ContractSpec() As IContractSpecifier
Set ContractSpec = mContractSpec
End Property

Public Property Let zIsBuyLeg(ByVal Value As Boolean)
Attribute zIsBuyLeg.VB_MemberFlags = "40"
mIsBuyLeg = Value
End Property

Public Property Get IsBuyLeg() As Boolean
IsBuyLeg = mIsBuyLeg
End Property

Public Property Let zRatio(ByVal Value As Long)
Attribute zRatio.VB_MemberFlags = "40"
mRatio = Value
End Property

Public Property Get Ratio() As Long
Ratio = mRatio
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pIsBuyLeg As Boolean, _
                ByVal pRatio As Long)
Set mContractSpec = pContractSpec
mIsBuyLeg = pIsBuyLeg
mRatio = pRatio
End Sub

Friend Function ToJSON() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToJSON"
On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "zContractSpec"
jb.EncodeValue mContractSpec

jb.EncodeName "zIsBuyLeg"
jb.EncodeValue mIsBuyLeg

jb.EncodeName "zRatio"
jb.EncodeValue mRatio

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToString() As String
Const ProcName As String = "ToString"

On Error GoTo Err

ToString = IIf(mIsBuyLeg, "BUY", "SELL") & ";" & _
            mRatio & ";" & _
            "(" & mContractSpec.ToString & ")"

Exit Function

Err:
ContractsGlobals.gHandleUnexpectedError ProcName, ModuleName

End Function

'@================================================================================
' Helper Functions
'@================================================================================


