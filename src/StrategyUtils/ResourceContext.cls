VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "ResourceContext"

'@================================================================================
' Member variables
'@================================================================================

Private mPrimaryTimeframe                           As ResourceIdentifier

Private mPrimaryBollingerBands                      As ResourceIdentifier
Private mPrimaryMACD                                As ResourceIdentifier
Private mPrimarySMA                                 As ResourceIdentifier
Private mPrimarySwing                               As ResourceIdentifier

Private mPrimaryBracketOrder                        As ResourceIdentifier

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

Friend Property Get PrimaryBollingerBands() As ResourceIdentifier
Set PrimaryBollingerBands = mPrimaryBollingerBands
End Property

Friend Property Get PrimaryBracketOrder() As ResourceIdentifier
Set PrimaryBracketOrder = mPrimaryBracketOrder
End Property

Friend Property Get PrimaryMACD() As ResourceIdentifier
Set PrimaryMACD = mPrimaryMACD
End Property

Friend Property Get PrimarySMA() As ResourceIdentifier
Set PrimarySMA = mPrimarySMA
End Property

Friend Property Get PrimarySwing() As ResourceIdentifier
Set PrimarySwing = mPrimarySwing
End Property

Friend Property Let PrimaryTimeframe(ByVal Value As ResourceIdentifier)
Set mPrimaryTimeframe = Value
End Property

Friend Property Get PrimaryTimeframe() As ResourceIdentifier
Set PrimaryTimeframe = mPrimaryTimeframe
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ClearPrimaryBracketOrder()
Set mPrimaryBracketOrder = Nothing
End Sub

Friend Function Clone() As ResourceContext
Set Clone = New ResourceContext
Clone.CloneFrom Me
End Function

Friend Sub CloneFrom(ByVal o As ResourceContext)
Set mPrimaryTimeframe = o.PrimaryTimeframe
Set mPrimaryBollingerBands = o.PrimaryBollingerBands
Set mPrimaryMACD = o.PrimaryMACD
Set mPrimarySMA = o.PrimarySMA
Set mPrimarySwing = o.PrimarySwing
End Sub

Friend Sub SetPrimaryBracketOrder(ByVal BracketOrder As ResourceIdentifier)
Set mPrimaryBracketOrder = BracketOrder
End Sub

Friend Sub SetPrimaryStudyOfType(ByVal pStudy As ResourceIdentifier)
Const ProcName As String = "SetPrimaryStudyOfType"
On Error GoTo Err

If TypeOf pStudy.Resource Is CmnStudiesLib27.BollingerBands Then
    If mPrimaryBollingerBands Is Nothing Then Set mPrimaryBollingerBands = pStudy
ElseIf TypeOf pStudy.Resource Is CmnStudiesLib27.MACD Then
    If mPrimaryMACD Is Nothing Then Set mPrimaryMACD = pStudy
ElseIf TypeOf pStudy.Resource Is CmnStudiesLib27.SMA Then
    If mPrimarySMA Is Nothing Then Set mPrimarySMA = pStudy
ElseIf TypeOf pStudy.Resource Is CmnStudiesLib27.Swing Then
    If mPrimarySwing Is Nothing Then Set mPrimarySwing = pStudy
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




