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

Public Property Get contractSpec() As ContractSpecifier
Set contractSpec = mContractSpec
End Property

Public Property Get isBuyLeg() As Boolean
isBuyLeg = mIsBuyLeg
End Property

Public Property Get ratio() As Long
ratio = mRatio
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function clone() As ComboLeg
Set clone = New ComboLeg
With clone
    .action = mAction
    .contractID = mContractID
    .exchange = mExchange
    .openClose = mOpenClose
    .ratio = mRatio
End With
End Function

Friend Sub initialise( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal isBuyLeg As Boolean, _
                ByVal ratio As Long)
Set mContractSpec = contractSpec
mIsBuyLeg = isBuyLeg
mRatio = ratio
End Sub

'Public Function ToString() As String
'
'ToString = "action=" & gOrderActionToString(mAction) & "; " & _
'            "contractid=" & mContractID & "; " & _
'            "exchange=" & mExchange & "; " & _
'            "openclose=" & mOpenClose & "; " & _
'            "ratio=" & mRatio & "; "
'
'End Function

'@================================================================================
' Helper Functions
'@================================================================================


