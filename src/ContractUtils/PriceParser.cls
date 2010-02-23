VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceParser"
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

Private Const ModuleName                            As String = "PriceParser"

'@================================================================================
' Member variables
'@================================================================================

Private mParsePriceAs32ndsPattern As String
Private mParsePriceAs32ndsAndFractionsPattern As String
Private mParsePriceAs64thsPattern As String
Private mParsePriceAs64thsAndFractionsPattern  As String
Private mParsePriceAsDecimalsPattern As String

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

Friend Property Let NumberOfDecimals( _
                ByVal value As Long)
mParsePriceAsDecimalsPattern = "^\d+($" & _
                            "|\.\d{1," & value & "}$)"

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub GenerateParsePriceAs32ndsPattern()
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

Public Sub GenerateParsePriceAs32ndsAndFractionsPattern()
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

Public Sub GenerateParsePriceAs64thsPattern()
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

Public Sub GenerateParsePriceAs64thsAndFractionsPattern()
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

Public Function ParsePriceAs32nds( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Dim lSubmatches As SubMatches

Const ProcName As String = "ParsePriceAs32nds"
Dim failpoint As String
On Error GoTo Err

If Not getSubmatches(pPriceString, mParsePriceAs32ndsPattern, lSubmatches) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 32

ParsePriceAs32nds = True
    
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs32ndsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Dim lSubmatches As SubMatches

Const ProcName As String = "ParsePriceAs32ndsAndFractions"
Dim failpoint As String
On Error GoTo Err

If Not getSubmatches(pPriceString, _
                    mParsePriceAs32ndsAndFractionsPattern, _
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

ParsePriceAs32ndsAndFractions = True
    
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs64ths( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Dim lSubmatches As SubMatches

Const ProcName As String = "ParsePriceAs64ths"
Dim failpoint As String
On Error GoTo Err

If Not getSubmatches(pPriceString, _
                    mParsePriceAs64thsPattern, _
                    lSubmatches) Then Exit Function

If Not gREgExp.Test(pPriceString) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 64

ParsePriceAs64ths = True
    
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAs64thsAndFractions( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean
Dim lSubmatches As SubMatches

Const ProcName As String = "ParsePriceAs64thsAndFractions"
Dim failpoint As String
On Error GoTo Err

If Not getSubmatches(pPriceString, _
                    mParsePriceAs64thsAndFractionsPattern, _
                    lSubmatches) Then Exit Function

If lSubmatches.Count = 0 Then Exit Function

pPrice = CDbl(lSubmatches(0))
If lSubmatches(3) <> "" Then pPrice = pPrice + CInt(lSubmatches(3)) / 64

If lSubmatches(4) <> "" Then
    If memberOf(lSubmatches(4), gHalfSixtyFourthIndicators) Then
        pPrice = pPrice + 1 / 128
    End If
End If

ParsePriceAs64thsAndFractions = True
    
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function ParsePriceAsDecimals( _
                ByVal pPriceString As String, _
                ByRef pPrice As Double) As Boolean

Const ProcName As String = "ParsePriceAsDecimals"
Dim failpoint As String
On Error GoTo Err

If IsMatched(pPriceString, mParsePriceAsDecimalsPattern) Then
    
    ' don't use CDBL here as we don't want to follow locale conventions (ie decimal point
    ' must be a period here)
    pPrice = Val(pPriceString)
    ParsePriceAsDecimals = True
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function IsMatched( _
                ByRef pPriceString As String, _
                ByRef pPattern As String) As Boolean
Dim lMatches As MatchCollection
Dim lMatch As Match

Const ProcName As String = "IsMatched"
Dim failpoint As String
On Error GoTo Err

gREgExp.Pattern = pPattern
IsMatched = gREgExp.Test(pPriceString)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

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

Private Function getSubmatches( _
                ByRef pPriceString As String, _
                ByRef pPattern As String, _
                ByRef pSubmatches As SubMatches) As Boolean
Dim lMatches As MatchCollection
Dim lMatch As Match

Const ProcName As String = "getSubmatches"
Dim failpoint As String
On Error GoTo Err

gREgExp.Pattern = pPattern
Set lMatches = gREgExp.Execute(pPriceString)

If lMatches.Count = 0 Then Exit Function

Set lMatch = lMatches(0)
Set pSubmatches = lMatch.SubMatches
    
getSubmatches = True
Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Private Function memberOf( _
                ByRef pInstring As String, _
                ByRef pChoices() As String) As Boolean
Dim i As Long
Const ProcName As String = "memberOf"
Dim failpoint As String
On Error GoTo Err

For i = 0 To UBound(pChoices)
    If pChoices(i) = pInstring Then
        memberOf = True
        Exit Function
    End If
Next

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

