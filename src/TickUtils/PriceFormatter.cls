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

Private Const ModuleName                            As String = "PriceFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mPriceFormatString As String

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

Friend Property Let PriceFormatString( _
                ByVal value As String)
mPriceFormatString = value
End Property
                
'@================================================================================
' Methods
'@================================================================================

Public Function FormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Dim fract As Double
Dim numerator As Long

Const ProcName As String = "FormatPriceAs32nds"
Dim failpoint As String
On Error GoTo Err

fract = pPrice - Int(pPrice)
numerator = fract * 32
FormatPriceAs32nds = Int(pPrice) & gDefaultThirtySecondsSeparator & Format(numerator, "00") & gDefaultThirtySecondsTerminator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Dim fract As Double
Dim numerator As Long
Dim priceString As String

Const ProcName As String = "FormatPriceAs32ndsAndFractions"
Dim failpoint As String
On Error GoTo Err

fract = pPrice - Int(pPrice)
numerator = fract * 128
priceString = Int(pPrice) & gDefaultThirtySecondsAndFractionsSeparator & Format(numerator \ 4, "00")
Select Case numerator Mod 4
Case 0
    priceString = priceString & gDefaultExactThirtySecondIndicator
Case 1
    priceString = priceString & gDefaultQuarterThirtySecondIndicator
Case 2
    priceString = priceString & gDefaultHalfThirtySecondIndicator
Case 3
    priceString = priceString & gDefaultThreeQuarterThirtySecondIndicator
End Select

FormatPriceAs32ndsAndFractions = priceString & gDefaultThirtySecondsAndFractionsTerminator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Dim fract As Double
Dim numerator As Long

Const ProcName As String = "FormatPriceAs64ths"
Dim failpoint As String
On Error GoTo Err

fract = pPrice - Int(pPrice)
numerator = fract * 64
FormatPriceAs64ths = Int(pPrice) & gDefaultSixtyFourthsSeparator & Format(numerator, "00") & gDefaultSixtyFourthsTerminator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Dim fract As Double
Dim numerator As Long
Dim priceString As String

Const ProcName As String = "FormatPriceAs64thsAndFractions"
Dim failpoint As String
On Error GoTo Err

fract = pPrice - Int(pPrice)
numerator = fract * 128
priceString = Int(pPrice) & gDefaultSixtyFourthsAndFractionsSeparator & Format(numerator \ 2, "00")
Select Case numerator Mod 2
Case 0
    priceString = priceString & gDefaultExactSixtyFourthIndicator
Case 1
    priceString = priceString & gDefaultHalfSixtyFourthIndicator
End Select

FormatPriceAs64thsAndFractions = priceString & gDefaultSixtyFourthsAndFractionsTerminator

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function FormatPriceAsDecimals( _
                ByVal pPrice As Double) As String
Const ProcName As String = "FormatPriceAsDecimals"
Dim failpoint As String
On Error GoTo Err

FormatPriceAsDecimals = Format(pPrice, mPriceFormatString)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================


