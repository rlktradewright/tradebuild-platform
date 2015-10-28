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

Private mBracketOrderMapping                        As New EnumerableCollection

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

Friend Function GetResourceIdForBracketOrder( _
                ByVal pBracketOrder As IBracketOrder) As ResourceIdentifier
Const ProcName As String = "GetResourceIdForBracketOrder"
On Error GoTo Err

Set GetResourceIdForBracketOrder = mBracketOrderMapping.Item(GetObjectKey(pBracketOrder))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub MapBracketOrderToResourceId( _
                ByVal pBracketOrder As IBracketOrder, _
                ByVal pIdentifer As ResourceIdentifier)
Const ProcName As String = "MapBracketOrderToResourceId"
On Error GoTo Err

mBracketOrderMapping.Add pIdentifer, GetObjectKey(pBracketOrder)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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




