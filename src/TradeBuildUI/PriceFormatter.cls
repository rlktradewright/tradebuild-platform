VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IPriceFormatter

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

Private Const ModuleName                            As String = "PriceFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mcontract As Contract

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IPriceFormatter Interface Members
'@================================================================================

Private Function IPriceFormatter_FormatPrice( _
                ByVal pValue As Double) As String
Const ProcName As String = "IPriceFormatter_FormatPrice"
On Error GoTo Err

IPriceFormatter_FormatPrice = mcontract.FormatPrice(pValue)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Property Let IPriceFormatter_IntegerYScale(ByVal RHS As Boolean)

End Property

Private Property Let IPriceFormatter_YScaleGridSpacing(ByVal RHS As Double)

End Property

Private Property Let IPriceFormatter_YScaleQuantum(ByVal RHS As Double)

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Contract(ByVal value As Contract)
Set mcontract = value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


