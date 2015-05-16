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

Private mSecType                                    As SecurityTypes
Private mTickSize                                   As Double

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

IPriceFormatter_FormatPrice = FormatPrice(pValue, mSecType, mTickSize)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double)
mSecType = pSecType
mTickSize = pTickSize
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


