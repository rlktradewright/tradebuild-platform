Attribute VB_Name = "GPriceParser"
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
 
Public Type TickSizePatternEntry
    TickSize            As Double
    Pattern             As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GPriceParser"

Public Const OneTenth                               As Double = 0.01
Public Const OneHalf                                As Double = 0.5
Public Const OneQuarter                             As Double = 0.25
Public Const OneEigth                               As Double = 0.125
Public Const OneSixteenth                           As Double = 0.0625
Public Const OneThirtySecond                        As Double = 0.03125
Public Const OneSixtyFourth                         As Double = 0.015625
Public Const OneHundredTwentyEighth                 As Double = 0.0078125

'@================================================================================
' Member variables
'@================================================================================

Private mParsePriceAs32ndsPattern As String
Private mParsePriceAs32ndsAndFractionsPattern As String
Private mParsePriceAs64thsPattern As String
Private mParsePriceAs64thsAndFractionsPattern  As String

Private mParsePriceAsDecimalsPatterns() As TickSizePatternEntry
Private mParsePriceAsDecimalsPatternsIndex As Long

Private mThirtySecondsSeparators() As String
Private mThirtySecondsTerminators() As String

Private mThirtySecondsAndFractionsSeparators() As String
Private mThirtySecondsAndFractionsTerminators() As String

Private mSixtyFourthsSeparators() As String
Private mSixtyFourthsTerminators() As String

Private mSixtyFourthsAndFractionsSeparators() As String
Private mSixtyFourthsAndFractionsTerminators() As String

Private mExactThirtySecondIndicators() As String
Private mQuarterThirtySecondIndicators() As String
Private mHalfThirtySecondIndicators() As String
Private mThreeQuarterThirtySecondIndicators() As String

Private mExactSixtyFourthIndicators() As String
Private mHalfSixtyFourthIndicators() As String

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

Public Property Get gDefaultExactSixtyFourthIndicator() As String
gDefaultExactSixtyFourthIndicator = mExactSixtyFourthIndicators(0)
End Property
                
Public Property Get gDefaultExactThirtySecondIndicator() As String
gDefaultExactThirtySecondIndicator = mExactThirtySecondIndicators(0)
End Property
                
Public Property Get gDefaultHalfSixtyFourthIndicator() As String
gDefaultHalfSixtyFourthIndicator = mHalfSixtyFourthIndicators(0)
End Property
                
Public Property Get gDefaultHalfThirtySecondIndicator() As String
gDefaultHalfThirtySecondIndicator = mHalfThirtySecondIndicators(0)
End Property
                
Public Property Get gDefaultQuarterThirtySecondIndicator() As String
gDefaultQuarterThirtySecondIndicator = mQuarterThirtySecondIndicators(0)
End Property
                
Public Property Get gDefaultSixtyFourthsAndFractionsSeparator() As String
gDefaultSixtyFourthsAndFractionsSeparator = mSixtyFourthsAndFractionsSeparators(0)
End Property
                
Public Property Get gDefaultSixtyFourthsAndFractionsTerminator() As String
gDefaultSixtyFourthsAndFractionsTerminator = mSixtyFourthsAndFractionsTerminators(0)
End Property
                
Public Property Get gDefaultSixtyFourthsSeparator() As String
gDefaultSixtyFourthsSeparator = mSixtyFourthsSeparators(0)
End Property
                
Public Property Get gDefaultSixtyFourthsTerminator() As String
gDefaultSixtyFourthsTerminator = mSixtyFourthsTerminators(0)
End Property
                
Public Property Get gDefaultThirtySecondsAndFractionsSeparator() As String
gDefaultThirtySecondsAndFractionsSeparator = mThirtySecondsAndFractionsSeparators(0)
End Property
                
Public Property Get gDefaultThirtySecondsAndFractionsTerminator() As String
gDefaultThirtySecondsAndFractionsTerminator = mThirtySecondsAndFractionsTerminators(0)
End Property
                
Public Property Get gDefaultThirtySecondsSeparator() As String
gDefaultThirtySecondsSeparator = mThirtySecondsSeparators(0)
End Property
                
Public Property Get gDefaultThirtySecondsTerminator() As String
gDefaultThirtySecondsTerminator = mThirtySecondsTerminators(0)
End Property
                
Public Property Get gDefaultThreeQuarterThirtySecondIndicator() As String
gDefaultThreeQuarterThirtySecondIndicator = mThreeQuarterThirtySecondIndicators(0)
End Property
                
Public Property Let gExactSixtyFourthIndicators( _
                ByRef Value() As String)
mExactSixtyFourthIndicators = Value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gExactSixtyFourthIndicators() As String()
gExactSixtyFourthIndicators = mExactSixtyFourthIndicators
End Property
                
Public Property Let gExactThirtySecondIndicators( _
                ByRef Value() As String)
mExactThirtySecondIndicators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gExactThirtySecondIndicators() As String()
gExactThirtySecondIndicators = mExactThirtySecondIndicators
End Property
                
Public Property Let gHalfSixtyFourthIndicators( _
                ByRef Value() As String)
mHalfSixtyFourthIndicators = Value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gHalfSixtyFourthIndicators() As String()
gHalfSixtyFourthIndicators = mHalfSixtyFourthIndicators
End Property
                
Public Property Let gHalfThirtySecondIndicators( _
                ByRef Value() As String)
mHalfThirtySecondIndicators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gHalfThirtySecondIndicators() As String()
gHalfThirtySecondIndicators = mHalfThirtySecondIndicators
End Property
                
Public Property Get gParsePriceAs32ndsPattern() As String
gParsePriceAs32ndsPattern = mParsePriceAs32ndsPattern
End Property

Public Property Get gParsePriceAs32ndsAndFractionsPattern() As String
gParsePriceAs32ndsAndFractionsPattern = mParsePriceAs32ndsAndFractionsPattern
End Property

Public Property Get gParsePriceAs64thsPattern() As String
gParsePriceAs64thsPattern = mParsePriceAs64thsPattern
End Property

Public Property Get gParsePriceAs64thsAndFractionsPattern() As String
gParsePriceAs64thsAndFractionsPattern = mParsePriceAs64thsAndFractionsPattern
End Property

Public Property Let gQuarterThirtySecondIndicators( _
                ByRef Value() As String)
mQuarterThirtySecondIndicators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gQuarterThirtySecondIndicators() As String()
gQuarterThirtySecondIndicators = mQuarterThirtySecondIndicators
End Property
                
Public Property Let gSixtyFourthsAndFractionsSeparators( _
                ByRef Value() As String)
mSixtyFourthsAndFractionsSeparators = Value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsSeparators() As String()
gSixtyFourthsAndFractionsSeparators = mSixtyFourthsAndFractionsSeparators
End Property
                
Public Property Let gSixtyFourthsAndFractionsTerminators( _
                ByRef Value() As String)
mSixtyFourthsAndFractionsTerminators = Value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsTerminators() As String()
gSixtyFourthsAndFractionsTerminators = mSixtyFourthsAndFractionsTerminators
End Property
                
Public Property Let gSixtyFourthsSeparators( _
                ByRef Value() As String)
mSixtyFourthsSeparators = Value
GPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsSeparators() As String()
gSixtyFourthsSeparators = mSixtyFourthsSeparators
End Property
                
Public Property Let gSixtyFourthsTerminators( _
                ByRef Value() As String)
mSixtyFourthsTerminators = Value
GPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsTerminators() As String()
gSixtyFourthsTerminators = mSixtyFourthsTerminators
End Property
                
Public Property Let gThirtySecondsAndFractionsSeparators( _
                ByRef Value() As String)
mThirtySecondsAndFractionsSeparators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsSeparators() As String()
gThirtySecondsAndFractionsSeparators = mThirtySecondsAndFractionsSeparators
End Property
                
Public Property Let gThirtySecondsAndFractionsTerminators( _
                ByRef Value() As String)
mThirtySecondsAndFractionsTerminators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsTerminators() As String()
gThirtySecondsAndFractionsTerminators = mThirtySecondsAndFractionsTerminators
End Property
                
Public Property Let gThirtySecondsSeparators( _
                ByRef Value() As String)
mThirtySecondsSeparators = Value
GPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsSeparators() As String()
gThirtySecondsSeparators = mThirtySecondsSeparators
End Property
                
Public Property Let gThirtySecondsTerminators( _
                ByRef Value() As String)
mThirtySecondsTerminators = Value
GPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsTerminators() As String()
gThirtySecondsTerminators = mThirtySecondsTerminators
End Property
                
Public Property Let gThreeQuarterThirtySecondIndicators( _
                ByRef Value() As String)
mThreeQuarterThirtySecondIndicators = Value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThreeQuarterThirtySecondIndicators() As String()
gThreeQuarterThirtySecondIndicators = mThreeQuarterThirtySecondIndicators
End Property
                
'@================================================================================
' Methods
'@================================================================================

Public Sub gInit()
mThirtySecondsSeparators = makeStringArray("'")
mThirtySecondsTerminators = makeStringArray(",/32")

mThirtySecondsAndFractionsSeparators = makeStringArray("'")
mThirtySecondsAndFractionsTerminators = makeStringArray(",/32")

mSixtyFourthsSeparators = makeStringArray(" ,''")
mSixtyFourthsTerminators = makeStringArray(",/64")

mSixtyFourthsAndFractionsSeparators = makeStringArray("''")
mSixtyFourthsAndFractionsTerminators = makeStringArray(",/64")

mExactThirtySecondIndicators = makeStringArray(",0")
mQuarterThirtySecondIndicators = makeStringArray("�,2")
mHalfThirtySecondIndicators = makeStringArray("+,5")
mThreeQuarterThirtySecondIndicators = makeStringArray("�,7")

mExactSixtyFourthIndicators = makeStringArray(",''")
mHalfSixtyFourthIndicators = makeStringArray("+,5")

GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
GPriceParser.GenerateParsePriceAs32ndsPattern
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
GPriceParser.GenerateParsePriceAs64thsPattern

ReDim mParsePriceAsDecimalsPatterns(7) As TickSizePatternEntry

End Sub

Public Function gParsePrice( _
                ByVal pPriceString As String, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePrice"
On Error GoTo Err

pPriceString = Trim$(pPriceString)

If pTickSize = -1 Then
    gParsePrice = gParsePriceGeneric(pPriceString, pPrice)
ElseIf pTickSize = OneThirtySecond Then
    gParsePrice = gParsePriceAs32nds(pPriceString, pPrice)
ElseIf pTickSize = OneSixtyFourth Then
    If pSecType = SecTypeFuture Then
        gParsePrice = gParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    Else
        gParsePrice = gParsePriceAs64ths(pPriceString, pPrice)
    End If
ElseIf pTickSize = OneHundredTwentyEighth Then
    If pSecType = SecTypeFuture Then
        gParsePrice = gParsePriceAs32ndsAndFractions(pPriceString, pPrice)
    Else
        gParsePrice = gParsePriceAs64thsAndFractions(pPriceString, pPrice)
    End If
Else
    gParsePrice = gParsePriceAsDecimals(pPriceString, pTickSize, pPrice)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAs32nds"
On Error GoTo Err

Dim lSubmatches As SubMatches
If Not getSubmatches(pPriceString, gParsePriceAs32ndsPattern, lSubmatches) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 32

gParsePriceAs32nds = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAs32ndsAndFractions"
On Error GoTo Err

Dim lSubmatches As SubMatches
If Not getSubmatches(pPriceString, _
                    gParsePriceAs32ndsAndFractionsPattern, _
                    lSubmatches) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 32

If lSubmatches(4) <> "" Then
    If memberOf(lSubmatches(4), gQuarterThirtySecondIndicators) Then
        pPrice = pPrice + 1 / 128
    ElseIf memberOf(lSubmatches(4), gHalfThirtySecondIndicators) Then
        pPrice = pPrice + 1 / 64
    ElseIf memberOf(lSubmatches(4), gThreeQuarterThirtySecondIndicators) Then
        pPrice = pPrice + 3 * 3 / 128
    End If
End If

gParsePriceAs32ndsAndFractions = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAs64ths"
On Error GoTo Err

Dim lSubmatches As SubMatches
If Not getSubmatches(pPriceString, _
                    gParsePriceAs64thsPattern, _
                    lSubmatches) Then Exit Function

If Not gRegExp.Test(pPriceString) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 64

gParsePriceAs64ths = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAs64thsAndFractions"
On Error GoTo Err

Dim lSubmatches As SubMatches
If Not getSubmatches(pPriceString, _
                    gParsePriceAs64thsAndFractionsPattern, _
                    lSubmatches) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 64

If lSubmatches(4) <> "" Then
    If memberOf(lSubmatches(4), gHalfSixtyFourthIndicators) Then
        pPrice = pPrice + 1 / 128
    End If
End If

gParsePriceAs64thsAndFractions = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceAsDecimals( _
                ByVal pPriceString As String, _
                ByVal pTickSize As Double, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAsDecimals"
On Error GoTo Err

If IsMatched(pPriceString, getParsePriceAsDecimalsPattern(pTickSize)) Then
    
    ' don't use CDBL here as we don't want to follow locale conventions (ie decimal point
    ' must be a period here)
    pPrice = Val(pPriceString)
    gParsePriceAsDecimals = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gParsePriceGeneric( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Const ProcName As String = "gParsePriceAsDecimals"
On Error GoTo Err

gParsePriceGeneric = True
If gParsePriceAs32nds(pPriceString, pPrice) Then
ElseIf gParsePriceAs32ndsAndFractions(pPriceString, pPrice) Then
ElseIf gParsePriceAs64ths(pPriceString, pPrice) Then
ElseIf gParsePriceAs64thsAndFractions(pPriceString, pPrice) Then
ElseIf gParsePriceAsDecimals(pPriceString, 0.00000001, pPrice) Then
Else
    gParsePriceGeneric = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addParsePriceAsDecimalsPattern(ByVal pTickSize As Double, pPattern As String)
Const ProcName As String = "addParsePriceAsDecimalsPattern"
On Error GoTo Err

If mParsePriceAsDecimalsPatternsIndex > UBound(mParsePriceAsDecimalsPatterns) Then ReDim Preserve mParsePriceAsDecimalsPatterns(2 * (UBound(mParsePriceAsDecimalsPatterns) + 1) - 1) As TickSizePatternEntry
mParsePriceAsDecimalsPatterns(mParsePriceAsDecimalsPatternsIndex).TickSize = pTickSize
mParsePriceAsDecimalsPatterns(mParsePriceAsDecimalsPatternsIndex).Pattern = pPattern

mParsePriceAsDecimalsPatternsIndex = mParsePriceAsDecimalsPatternsIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function convertToRegexpChoices( _
                ByRef choiceStrings() As String) As String
Dim s As String

Dim i As Long
For i = 0 To UBound(choiceStrings)
    Dim choiceString As String: choiceString = choiceStrings(i)
    If i <> 0 Then s = s & "|"
    s = s & escapeRegexSpecialChar(choiceString)
Next

convertToRegexpChoices = s
End Function

Private Function escapeRegexSpecialChar( _
                ByRef inString As String) As String
Dim s As String

Dim i As Long
For i = 1 To Len(inString)
    Dim ch As String: ch = Mid$(inString, i, 1)
    Select Case ch
    Case "*", "+", "?", "^", "$", "[", "]", "{", "}", "(", ")", "|", "/", "\"
        s = s & "\" & ch
    Case Else
        s = s & ch
    End Select
Next

escapeRegexSpecialChar = s
End Function

Private Sub GenerateParsePriceAs32ndsPattern()
mParsePriceAs32ndsPattern = _
                            "^(\d+)" & _
                            "(" & _
                                "($" & _
                                    "|" & convertToRegexpChoices(gThirtySecondsSeparators) & _
                                ")" & _
                                "([0-2][0-9]|30|31)" & _
                                "(" & _
                                      convertToRegexpChoices(gThirtySecondsTerminators) & _
                                ")$" & _
                            ")"
End Sub

Private Sub GenerateParsePriceAs32ndsAndFractionsPattern()
mParsePriceAs32ndsAndFractionsPattern = _
                            "^(\d+)" & _
                            "(" & _
                                "($" & _
                                    "|" & convertToRegexpChoices(gThirtySecondsAndFractionsSeparators) & _
                                ")" & _
                                "([0-2][0-9]|30|31)" & _
                                "(" & _
                                    convertToRegexpChoices(gExactThirtySecondIndicators) & _
                                    "|" & convertToRegexpChoices(gQuarterThirtySecondIndicators) & _
                                    "|" & convertToRegexpChoices(gHalfThirtySecondIndicators) & _
                                    "|" & convertToRegexpChoices(gThreeQuarterThirtySecondIndicators) & _
                                ")" & _
                                "(" & _
                                    convertToRegexpChoices(gThirtySecondsAndFractionsTerminators) & _
                                ")$" & _
                            ")"
End Sub

Private Sub GenerateParsePriceAs64thsPattern()
mParsePriceAs64thsPattern = _
                            "^(\d+)" & _
                            "(" & _
                                "($" & _
                                    "|" & convertToRegexpChoices(gSixtyFourthsSeparators) & _
                                ")" & _
                                "([0-5][0-9]|60|61|62|63)" & _
                                "(" & _
                                    convertToRegexpChoices(gSixtyFourthsTerminators) & _
                                ")$" & _
                            ")"
End Sub

Private Sub GenerateParsePriceAs64thsAndFractionsPattern()
mParsePriceAs64thsAndFractionsPattern = _
                            "^(\d+)" & _
                            "(" & _
                                "($" & _
                                    "|" & convertToRegexpChoices(gSixtyFourthsAndFractionsSeparators) & _
                                ")" & _
                                "([0-5][0-9]|60|61|62|63)" & _
                                "(" & _
                                    convertToRegexpChoices(gExactSixtyFourthIndicators) & _
                                    "|" & convertToRegexpChoices(gHalfSixtyFourthIndicators) & _
                                ")" & _
                                "(" & _
                                    "|" & convertToRegexpChoices(gSixtyFourthsAndFractionsTerminators) & _
                                ")$" & _
                            ")"
End Sub

Private Function generateParsePriceAsDecimalsPattern(ByVal pTickSize As Double) As String
Const ProcName As String = "generateParsePriceAsDecimalsPattern"
On Error GoTo Err

If pTickSize = 0# Then pTickSize = 0.00000000000001

Dim minTickString As String: minTickString = Format(pTickSize, "0.##############")

Dim lNumberOfDecimals As Long: lNumberOfDecimals = Len(minTickString) - 2

generateParsePriceAsDecimalsPattern = "^\d+($" & _
                            "|\.\d{1," & lNumberOfDecimals & "}$)"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getParsePriceAsDecimalsPattern(ByVal pTickSize As Double) As String
Const ProcName As String = "getParsePriceAsDecimalsPattern"
On Error GoTo Err

Dim lPattern As String

Dim i As Long
For i = 0 To mParsePriceAsDecimalsPatternsIndex - 1
    If mParsePriceAsDecimalsPatterns(i).TickSize = pTickSize Then
        getParsePriceAsDecimalsPattern = mParsePriceAsDecimalsPatterns(i).Pattern
        Exit Function
    End If
Next

lPattern = generateParsePriceAsDecimalsPattern(pTickSize)
addParsePriceAsDecimalsPattern pTickSize, lPattern
getParsePriceAsDecimalsPattern = lPattern

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSubmatches( _
                ByRef pPriceString As String, _
                ByRef pPattern As String, _
                ByRef pSubmatches As SubMatches) As Boolean
Const ProcName As String = "getSubmatches"
On Error GoTo Err

gRegExp.Pattern = pPattern
Dim lMatches As MatchCollection: Set lMatches = gRegExp.Execute(pPriceString)

If lMatches.Count = 0 Then Exit Function

Dim lMatch As Match: Set lMatch = lMatches(0)
Set pSubmatches = lMatch.SubMatches
    
getSubmatches = True
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IsMatched( _
                ByRef pPriceString As String, _
                ByRef pPattern As String) As Boolean
Const ProcName As String = "IsMatched"
On Error GoTo Err

gRegExp.Pattern = pPattern
IsMatched = gRegExp.Test(pPriceString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function makeStringArray(ByRef inString) As String()
makeStringArray = Split(inString, ",")
End Function

Private Function memberOf( _
                ByRef pInstring As String, _
                ByRef pChoices() As String) As Boolean
Const ProcName As String = "memberOf"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(pChoices)
    If pChoices(i) = pInstring Then
        memberOf = True
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




