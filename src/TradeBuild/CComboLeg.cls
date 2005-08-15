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
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

Implements TradeBuildSP.IComboLeg

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mContractID As Long
Private mAction As String
Private mRatio As Long
Private mExchange As String
Private mOpenClose As LegOpenClose

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' IComboLeg Interface Members
'================================================================================

Private Property Let IComboLeg_Action(ByVal RHS As String)
action = RHS
End Property

Private Property Get IComboLeg_Action() As String
IComboLeg_Action = action
End Property

Private Function IComboLeg_Clone() As TradeBuildSP.IComboLeg
Set IComboLeg_Clone = Clone
End Function

Private Property Let IComboLeg_ContractID(ByVal RHS As Long)
contractID = RHS
End Property

Private Property Get IComboLeg_ContractID() As Long
IComboLeg_ContractID = contractID
End Property

Private Property Let IComboLeg_Exchange(ByVal RHS As String)
exchange = RHS
End Property

Private Property Get IComboLeg_Exchange() As String
IComboLeg_Exchange = exchange
End Property

Private Property Let IComboLeg_OpenClose(ByVal RHS As TradeBuildSP.LegOpenClose)
openClose = RHS
End Property

Private Property Get IComboLeg_OpenClose() As TradeBuildSP.LegOpenClose
IComboLeg_OpenClose = openClose
End Property

Private Property Let IComboLeg_Ratio(ByVal RHS As Long)
ratio = RHS
End Property

Private Property Get IComboLeg_Ratio() As Long
IComboLeg_Ratio = ratio
End Property

Private Function IComboLeg_ToString() As String
IComboLeg_ToString = ToString
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get action() As String
action = mAction
End Property

Public Property Let action(ByVal value As String)
mAction = UCase$(value)
End Property

Public Property Get contractID() As Long
contractID = mContractID
End Property

Public Property Let contractID(ByVal value As Long)
mContractID = value
End Property

Public Property Get exchange() As String
exchange = mExchange
End Property

Public Property Let exchange(ByVal value As String)
mExchange = UCase$(value)
End Property

Public Property Let openClose(ByVal value As LegOpenClose)
mOpenClose = value
End Property

Public Property Get openClose() As LegOpenClose
openClose = mOpenClose
End Property

Public Property Get ratio() As Long
ratio = mRatio
End Property

Public Property Let ratio(ByVal value As Long)
mRatio = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function Clone() As ComboLeg
Set Clone = New ComboLeg
With Clone
    .action = mAction
    .contractID = mContractID
    .exchange = mExchange
    .openClose = mOpenClose
    .ratio = mRatio
End With
End Function

Public Function ToString() As String

ToString = "action=" & mAction & vbCrLf & _
            vbTab & "contractid=" & mContractID & vbCrLf & _
            vbTab & "exchange=" & mExchange & vbCrLf & _
            vbTab & "openclose=" & mOpenClose & vbCrLf & _
            vbTab & "ratio=" & mRatio & vbCrLf

End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub Class_ReadProperties(PropBag As PropertyBag)
mContractID = PropBag.ReadProperty("conid")
mAction = PropBag.ReadProperty("action")
mRatio = PropBag.ReadProperty("ratio")
mExchange = PropBag.ReadProperty("exchange")
mOpenClose = PropBag.ReadProperty("openclose")
End Sub

Private Sub Class_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "conid", mContractID
PropBag.WriteProperty "action", mAction
PropBag.WriteProperty "ratio", mRatio
PropBag.WriteProperty "exchange", mExchange
PropBag.ReadProperty "openclose", mOpenClose
End Sub

