Attribute VB_Name = "GPriceFormatter"
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

Private Const ModuleName                            As String = "GPriceFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mPriceFormatStrings()                       As TickSizePatternEntry
Private mPriceFormatStringsIndex                    As Long

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

'@================================================================================
' Methods
'@================================================================================

Public Function gFormatPriceAs32nds( _
                ByVal pPrice As Double) As String
Const ProcName As String = "gFormatPriceAs32nds"
On Error GoTo Err

Dim fract As Double
Dim numerator As Long

fract = pPrice - Int(pPrice)
numerator = fract * 32
gFormatPriceAs32nds = Int(pPrice) & gDefaultThirtySecondsSeparator & Format(numerator, "00") & gDefaultThirtySecondsTerminator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFormatPriceAs32ndsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "gFormatPriceAs32ndsAndFractions"
On Error GoTo Err

Dim fract As Double
Dim numerator As Long
Dim priceString As String

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

gFormatPriceAs32ndsAndFractions = priceString & gDefaultThirtySecondsAndFractionsTerminator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFormatPriceAs64ths( _
                ByVal pPrice As Double) As String
Const ProcName As String = "gFormatPriceAs64ths"
On Error GoTo Err

Dim fract As Double
Dim numerator As Long

fract = pPrice - Int(pPrice)
numerator = fract * 64
gFormatPriceAs64ths = Int(pPrice) & gDefaultSixtyFourthsSeparator & Format(numerator, "00") & gDefaultSixtyFourthsTerminator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFormatPriceAs64thsAndFractions( _
                ByVal pPrice As Double) As String
Const ProcName As String = "gFormatPriceAs64thsAndFractions"
On Error GoTo Err

Dim fract As Double
Dim numerator As Long
Dim priceString As String

fract = pPrice - Int(pPrice)
numerator = fract * 128
priceString = Int(pPrice) & gDefaultSixtyFourthsAndFractionsSeparator & Format(numerator \ 2, "00")
Select Case numerator Mod 2
Case 0
    priceString = priceString & gDefaultExactSixtyFourthIndicator
Case 1
    priceString = priceString & gDefaultHalfSixtyFourthIndicator
End Select

gFormatPriceAs64thsAndFractions = priceString & gDefaultSixtyFourthsAndFractionsTerminator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFormatPriceAsDecimals( _
                ByVal pPrice As Double, _
                ByVal pTickSize As Double) As String
Const ProcName As String = "gFormatPriceAsDecimals"
Dim failpoint As String
On Error GoTo Err

gFormatPriceAsDecimals = Format(pPrice, getPriceFormatString(pTickSize))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInit()
ReDim mPriceFormatStrings(7) As TickSizePatternEntry
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addPriceFormatString(ByVal pTickSize As Double, pFormatString As String)
Const ProcName As String = "addPriceFormatString"
On Error GoTo Err

If mPriceFormatStringsIndex > UBound(mPriceFormatStrings) Then ReDim Preserve mPriceFormatStrings(2 * (UBound(mPriceFormatStrings) + 1) - 1) As TickSizePatternEntry
mPriceFormatStrings(mPriceFormatStringsIndex).TickSize = pTickSize
mPriceFormatStrings(mPriceFormatStringsIndex).Pattern = pFormatString

mPriceFormatStringsIndex = mPriceFormatStringsIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function generatePriceFormatString(ByVal pTickSize As Double)
Const ProcName As String = "generatePriceFormatString"
On Error GoTo Err

Dim minTickString As String
Dim lNumberOfDecimals As Long

minTickString = Format(pTickSize, "0.##############")

lNumberOfDecimals = Len(minTickString) - 2

If lNumberOfDecimals = 0 Then
    generatePriceFormatString = "0"
Else
    generatePriceFormatString = "0." & String(lNumberOfDecimals, "0")
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPriceFormatString(ByVal pTickSize As Double) As String
Const ProcName As String = "getPriceFormatString"
On Error GoTo Err

Dim i As Long
Dim lPattern As String

For i = 0 To mPriceFormatStringsIndex - 1
    If mPriceFormatStrings(i).TickSize = pTickSize Then
        getPriceFormatString = mPriceFormatStrings(i).Pattern
        Exit Function
    End If
Next

lPattern = generatePriceFormatString(pTickSize)
addPriceFormatString pTickSize, lPattern
getPriceFormatString = lPattern

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




