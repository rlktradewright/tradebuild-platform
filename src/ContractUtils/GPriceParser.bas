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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GPriceParser"

'@================================================================================
' Member variables
'@================================================================================

Private mParsePriceAs32ndsPattern As String
Private mParsePriceAs32ndsAndFractionsPattern As String
Private mParsePriceAs64thsPattern As String
Private mParsePriceAs64thsAndFractionsPattern  As String

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
                ByRef value() As String)
mExactSixtyFourthIndicators = value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gExactSixtyFourthIndicators() As String()
gExactSixtyFourthIndicators = mExactSixtyFourthIndicators
End Property
                
Public Property Let gExactThirtySecondIndicators( _
                ByRef value() As String)
mExactThirtySecondIndicators = value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gExactThirtySecondIndicators() As String()
gExactThirtySecondIndicators = mExactThirtySecondIndicators
End Property
                
Public Property Let gHalfSixtyFourthIndicators( _
                ByRef value() As String)
mHalfSixtyFourthIndicators = value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gHalfSixtyFourthIndicators() As String()
gHalfSixtyFourthIndicators = mHalfSixtyFourthIndicators
End Property
                
Public Property Let gHalfThirtySecondIndicators( _
                ByRef value() As String)
mHalfThirtySecondIndicators = value
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
                ByRef value() As String)
mQuarterThirtySecondIndicators = value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gQuarterThirtySecondIndicators() As String()
gQuarterThirtySecondIndicators = mQuarterThirtySecondIndicators
End Property
                
Public Property Let gSixtyFourthsAndFractionsSeparators( _
                ByRef value() As String)
mSixtyFourthsAndFractionsSeparators = value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsSeparators() As String()
gSixtyFourthsAndFractionsSeparators = mSixtyFourthsAndFractionsSeparators
End Property
                
Public Property Let gSixtyFourthsAndFractionsTerminators( _
                ByRef value() As String)
mSixtyFourthsAndFractionsTerminators = value
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsTerminators() As String()
gSixtyFourthsAndFractionsTerminators = mSixtyFourthsAndFractionsTerminators
End Property
                
Public Property Let gSixtyFourthsSeparators( _
                ByRef value() As String)
mSixtyFourthsSeparators = value
GPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsSeparators() As String()
gSixtyFourthsSeparators = mSixtyFourthsSeparators
End Property
                
Public Property Let gSixtyFourthsTerminators( _
                ByRef value() As String)
mSixtyFourthsTerminators = value
GPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsTerminators() As String()
gSixtyFourthsTerminators = mSixtyFourthsTerminators
End Property
                
Public Property Let gThirtySecondsAndFractionsSeparators( _
                ByRef value() As String)
mThirtySecondsAndFractionsSeparators = value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsSeparators() As String()
gThirtySecondsAndFractionsSeparators = mThirtySecondsAndFractionsSeparators
End Property
                
Public Property Let gThirtySecondsAndFractionsTerminators( _
                ByRef value() As String)
mThirtySecondsAndFractionsTerminators = value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsTerminators() As String()
gThirtySecondsAndFractionsTerminators = mThirtySecondsAndFractionsTerminators
End Property
                
Public Property Let gThirtySecondsSeparators( _
                ByRef value() As String)
mThirtySecondsSeparators = value
GPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsSeparators() As String()
gThirtySecondsSeparators = mThirtySecondsSeparators
End Property
                
Public Property Let gThirtySecondsTerminators( _
                ByRef value() As String)
mThirtySecondsTerminators = value
GPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsTerminators() As String()
gThirtySecondsTerminators = mThirtySecondsTerminators
End Property
                
Public Property Let gThreeQuarterThirtySecondIndicators( _
                ByRef value() As String)
mThreeQuarterThirtySecondIndicators = value
GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThreeQuarterThirtySecondIndicators() As String()
gThreeQuarterThirtySecondIndicators = mThreeQuarterThirtySecondIndicators
End Property
                
'@================================================================================
' Methods
'@================================================================================

Public Sub Main()
mThirtySecondsSeparators = makeStringArray("'")
mThirtySecondsTerminators = makeStringArray(",/32")

mThirtySecondsAndFractionsSeparators = makeStringArray("'")
mThirtySecondsAndFractionsTerminators = makeStringArray(",/32")

mSixtyFourthsSeparators = makeStringArray(" ,''")
mSixtyFourthsTerminators = makeStringArray(",/64")

mSixtyFourthsAndFractionsSeparators = makeStringArray("''")
mSixtyFourthsAndFractionsTerminators = makeStringArray(",/64")

mExactThirtySecondIndicators = makeStringArray(",0")
mQuarterThirtySecondIndicators = makeStringArray("¼,2")
mHalfThirtySecondIndicators = makeStringArray("+,5")
mThreeQuarterThirtySecondIndicators = makeStringArray("¾,7")

mExactSixtyFourthIndicators = makeStringArray(",''")
mHalfSixtyFourthIndicators = makeStringArray("+,5")

GPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
GPriceParser.GenerateParsePriceAs32ndsPattern
GPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
GPriceParser.GenerateParsePriceAs64thsPattern

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function convertToRegexpChoices( _
                ByRef choiceStrings() As String) As String
Dim s As String
Dim i As Long
Dim choiceString As String

For i = 0 To UBound(choiceStrings)
    choiceString = choiceStrings(i)
    If i <> 0 Then s = s & "|"
    s = s & escapeRegexSpecialChar(choiceString)
Next

convertToRegexpChoices = s
End Function

Private Function escapeRegexSpecialChar( _
                ByRef inString As String) As String
Dim i As Long
Dim s As String
Dim ch As String

For i = 1 To Len(inString)
    ch = Mid$(inString, i, 1)
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

Private Function makeStringArray(ByRef inString) As String()
makeStringArray = Split(inString, ",")
End Function


