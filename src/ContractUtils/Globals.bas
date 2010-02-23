Attribute VB_Name = "Globals"
Option Explicit

''
' Description here
'
' @remarks
' @see
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

Public Const ProjectName                                As String = "ContractUtils26"
Private Const ModuleName                                As String = "Globals"

Public Const ConfigSectionContractSpecifier             As String = "Specifier"

Public Const ConfigSettingContractSpecCurrency          As String = ConfigSectionContractSpecifier & "&Currency"
Public Const ConfigSettingContractSpecExpiry            As String = ConfigSectionContractSpecifier & "&Expiry"
Public Const ConfigSettingContractSpecExchange          As String = ConfigSectionContractSpecifier & "&Exchange"
Public Const ConfigSettingContractSpecLocalSYmbol       As String = ConfigSectionContractSpecifier & "&LocalSymbol"
Public Const ConfigSettingContractSpecRight             As String = ConfigSectionContractSpecifier & "&Right"
Public Const ConfigSettingContractSpecSecType           As String = ConfigSectionContractSpecifier & "&SecType"
Public Const ConfigSettingContractSpecStrikePrice       As String = ConfigSectionContractSpecifier & "&StrikePrice"
Public Const ConfigSettingContractSpecSymbol            As String = ConfigSectionContractSpecifier & "&Symbol"

Public Const ConfigSettingDaysBeforeExpiryToSwitch      As String = "&DaysBeforeExpiryToSwitch"
Public Const ConfigSettingDescription                   As String = "&Description"
Public Const ConfigSettingExpiryDate                    As String = "&ExpiryDate"
Public Const ConfigSettingMultiplier                    As String = "&Multiplier"
Public Const ConfigSettingSessionEndTime                As String = "&SessionEndTime"
Public Const ConfigSettingSessionStartTime              As String = "&SessionStartTime"
Public Const ConfigSettingSessionTickSize               As String = "&TickSize"
Public Const ConfigSettingSessionTimezone               As String = "&Timezone"

'@================================================================================
' Member variables
'@================================================================================

Private mExchangeCodes() As String
Private mMaxExchangeCodesIndex As Long

Private mCurrencyDescs() As CurrencyDescriptor
Private mMaxCurrencyDescsIndex As Long

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
gPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gExactSixtyFourthIndicators() As String()
gExactSixtyFourthIndicators = mExactSixtyFourthIndicators
End Property
                
Public Property Let gExactThirtySecondIndicators( _
                ByRef value() As String)
mExactThirtySecondIndicators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gExactThirtySecondIndicators() As String()
gExactThirtySecondIndicators = mExactThirtySecondIndicators
End Property
                
Public Property Let gHalfSixtyFourthIndicators( _
                ByRef value() As String)
mHalfSixtyFourthIndicators = value
gPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gHalfSixtyFourthIndicators() As String()
gHalfSixtyFourthIndicators = mHalfSixtyFourthIndicators
End Property
                
Public Property Let gHalfThirtySecondIndicators( _
                ByRef value() As String)
mHalfThirtySecondIndicators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gHalfThirtySecondIndicators() As String()
gHalfThirtySecondIndicators = mHalfThirtySecondIndicators
End Property
                
Public Property Get gPriceFormatter() As PriceFormatter
Static lPriceFormatter As PriceFormatter
If lPriceFormatter Is Nothing Then Set lPriceFormatter = New PriceFormatter
Set gPriceFormatter = lPriceFormatter
End Property

Public Property Get gPriceParser() As PriceParser
Static lPriceParser As PriceParser
If lPriceParser Is Nothing Then Set lPriceParser = New PriceParser
Set gPriceParser = lPriceParser
End Property

Public Property Let gQuarterThirtySecondIndicators( _
                ByRef value() As String)
mQuarterThirtySecondIndicators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gQuarterThirtySecondIndicators() As String()
gQuarterThirtySecondIndicators = mQuarterThirtySecondIndicators
End Property
                
Public Property Get gREgExp() As RegExp
Static lRegExp As RegExp
If lRegExp Is Nothing Then Set lRegExp = New RegExp
Set gREgExp = lRegExp
End Property

Public Property Let gSixtyFourthsAndFractionsSeparators( _
                ByRef value() As String)
mSixtyFourthsAndFractionsSeparators = value
gPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsSeparators() As String()
gSixtyFourthsAndFractionsSeparators = mSixtyFourthsAndFractionsSeparators
End Property
                
Public Property Let gSixtyFourthsAndFractionsTerminators( _
                ByRef value() As String)
mSixtyFourthsAndFractionsTerminators = value
gPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
End Property
                
Public Property Get gSixtyFourthsAndFractionsTerminators() As String()
gSixtyFourthsAndFractionsTerminators = mSixtyFourthsAndFractionsTerminators
End Property
                
Public Property Let gSixtyFourthsSeparators( _
                ByRef value() As String)
mSixtyFourthsSeparators = value
gPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsSeparators() As String()
gSixtyFourthsSeparators = mSixtyFourthsSeparators
End Property
                
Public Property Let gSixtyFourthsTerminators( _
                ByRef value() As String)
mSixtyFourthsTerminators = value
gPriceParser.GenerateParsePriceAs64thsPattern
End Property
                
Public Property Get gSixtyFourthsTerminators() As String()
gSixtyFourthsTerminators = mSixtyFourthsTerminators
End Property
                
Public Property Let gThirtySecondsAndFractionsSeparators( _
                ByRef value() As String)
mThirtySecondsAndFractionsSeparators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsSeparators() As String()
gThirtySecondsAndFractionsSeparators = mThirtySecondsAndFractionsSeparators
End Property
                
Public Property Let gThirtySecondsAndFractionsTerminators( _
                ByRef value() As String)
mThirtySecondsAndFractionsTerminators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThirtySecondsAndFractionsTerminators() As String()
gThirtySecondsAndFractionsTerminators = mThirtySecondsAndFractionsTerminators
End Property
                
Public Property Let gThirtySecondsSeparators( _
                ByRef value() As String)
mThirtySecondsSeparators = value
gPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsSeparators() As String()
gThirtySecondsSeparators = mThirtySecondsSeparators
End Property
                
Public Property Let gThirtySecondsTerminators( _
                ByRef value() As String)
mThirtySecondsTerminators = value
gPriceParser.GenerateParsePriceAs32ndsPattern
End Property
                
Public Property Get gThirtySecondsTerminators() As String()
gThirtySecondsTerminators = mThirtySecondsTerminators
End Property
                
Public Property Let gThreeQuarterThirtySecondIndicators( _
                ByRef value() As String)
mThreeQuarterThirtySecondIndicators = value
gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
End Property
                
Public Property Get gThreeQuarterThirtySecondIndicators() As String()
gThreeQuarterThirtySecondIndicators = mThreeQuarterThirtySecondIndicators
End Property
                
'@================================================================================
' Methods
'@================================================================================

Public Function gCreateContractSpecifier( _
                Optional ByVal LocalSymbol As String, _
                Optional ByVal Symbol As String, _
                Optional ByVal Exchange As String, _
                Optional ByVal SecType As SecurityTypes = SecTypeNone, _
                Optional ByVal CurrencyCode As String, _
                Optional ByVal Expiry As String, _
                Optional ByVal Strike As Double, _
                Optional ByVal Right As OptionRights = OptNone) As ContractSpecifier
Const ProcName As String = "gCreateContractSpecifier"
Dim failpoint As String
On Error GoTo Err

If LocalSymbol = "" And Symbol = "" Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Symbol must be supplied if LocalSymbol is not supplied"
End If

If Exchange <> "" And _
    Not gIsValidExchangeCode(Exchange) _
Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "'" & Exchange & "' is not a valid Exchange code"
End If

If Expiry <> "" Then
    If Not gIsValidExpiry(Expiry) Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "'" & Expiry & "' is not a valid Expiry format"
    End If
End If

Select Case SecType
Case 0  ' ie not supplied
Case SecTypeStock
Case SecTypeFuture
Case SecTypeOption, SecTypeFuturesOption
    If Strike < 0 Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Strike must be > 0"
    End If
    Select Case Right
    Case OptCall
    Case OptPut
    Case OptNone
    Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "'" & Right & "' is not a valid option Right"
    End Select
Case SecTypeCash
Case SecTypeCombo
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Sectype 'combo' is not permissible"
Case SecTypeIndex
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "'" & SecType & "' is not a valid secType"
End Select

Set gCreateContractSpecifier = New ContractSpecifier
With gCreateContractSpecifier
    .LocalSymbol = LocalSymbol
    .Symbol = Symbol
    .Exchange = Exchange
    .SecType = SecType
    .CurrencyCode = CurrencyCode
    .Expiry = Expiry
    .Strike = Strike
    .Right = Right
End With

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gGetCurrencyDescriptor( _
                ByVal code As String) As CurrencyDescriptor
Dim index As Long
Const ProcName As String = "gGetCurrencyDescriptor"
Dim failpoint As String
On Error GoTo Err

If mMaxCurrencyDescsIndex = 0 Then setupCurrencyDescs
index = getCurrencyIndex(code)
If index < 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Invalid currency code"
End If

gGetCurrencyDescriptor = mCurrencyDescs(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gGetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "gGetCurrencyDescriptors"
Dim failpoint As String
On Error GoTo Err

If mMaxCurrencyDescsIndex = 0 Then setupCurrencyDescs
gGetCurrencyDescriptors = mCurrencyDescs

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gGetExchangeCodes() As String()
Const ProcName As String = "gGetExchangeCodes"
Dim failpoint As String
On Error GoTo Err

If mMaxExchangeCodesIndex = 0 Then setupExchangeCodes
gGetExchangeCodes = mExchangeCodes

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gIsValidCurrencyCode(ByVal code As String) As Boolean
Const ProcName As String = "gIsValidCurrencyCode"
Dim failpoint As String
On Error GoTo Err

gIsValidCurrencyCode = (getCurrencyIndex(code) >= 0)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gIsValidExchangeCode(ByVal code As String) As Boolean
Dim bottom As Long
Dim top As Long
Dim middle As Long

Const ProcName As String = "gIsValidExchangeCode"
Dim failpoint As String
On Error GoTo Err

If mMaxExchangeCodesIndex = 0 Then setupExchangeCodes

code = UCase$(code)
bottom = 0
top = mMaxExchangeCodesIndex
middle = Fix((bottom + top) / 2)

Do
    If code < mExchangeCodes(middle) Then
        top = middle
    ElseIf code > mExchangeCodes(middle) Then
        bottom = middle
    Else
        gIsValidExchangeCode = True
        Exit Function
    End If
    middle = Fix((bottom + top) / 2)
Loop Until bottom = middle

If code = mExchangeCodes(middle) Then gIsValidExchangeCode = True

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gIsValidExpiry( _
                ByVal value As String) As Boolean
Dim d As Date

Const ProcName As String = "gIsValidExpiry"
Dim failpoint As String
On Error GoTo Err

If IsDate(value) Then
    d = CDate(value)
ElseIf Len(value) = 8 Then
    Dim datestring As String
    datestring = Left$(value, 4) & "/" & Mid$(value, 5, 2) & "/" & Right$(value, 2)
    If IsDate(datestring) Then d = CDate(datestring)
End If

If d <> 0 Then
    If d >= CDate((Year(Now) - 20) & "/01/01") And d <= CDate((Year(Now) + 10) & "/12/31") Then
        gIsValidExpiry = True
        Exit Function
    End If
End If

If Len(value) = 6 Then
    If IsInteger(value, (Year(Now) - 20) * 100 + 1, (Year(Now) + 10) * 100 + 12) Then
        If Right$(value, 2) <= 12 Then
            gIsValidExpiry = True
            Exit Function
        End If
    End If
End If

gIsValidExpiry = False

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function gIsValidSecType( _
                ByVal value As Long) As Boolean
gIsValidSecType = True
Select Case value
Case SecTypeStock
Case SecTypeFuture
Case SecTypeOption
Case SecTypeFuturesOption
Case SecTypeCash
Case SecTypeCombo
Case SecTypeIndex
Case Else
    gIsValidSecType = False
End Select
End Function

Public Function gOptionRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    gOptionRightFromString = OptNone
Case "CALL", "C"
    gOptionRightFromString = OptCall
Case "PUT", "P"
    gOptionRightFromString = OptPut
End Select
End Function

Public Function gOptionRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    gOptionRightToString = ""
Case OptCall
    gOptionRightToString = "Call"
Case OptPut
    gOptionRightToString = "Put"
End Select
End Function

Public Function gSecTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STOCK", "STK"
    gSecTypeFromString = SecTypeStock
Case "FUTURE", "FUT"
    gSecTypeFromString = SecTypeFuture
Case "OPTION", "OPT"
    gSecTypeFromString = SecTypeOption
Case "FUTURES OPTION", "FOP"
    gSecTypeFromString = SecTypeFuturesOption
Case "CASH"
    gSecTypeFromString = SecTypeCash
Case "COMBO", "CMB"
    gSecTypeFromString = SecTypeCombo
Case "INDEX", "IND"
    gSecTypeFromString = SecTypeIndex
End Select
End Function

Public Function gSecTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToString = "Stock"
Case SecTypeFuture
    gSecTypeToString = "Future"
Case SecTypeOption
    gSecTypeToString = "Option"
Case SecTypeFuturesOption
    gSecTypeToString = "Futures Option"
Case SecTypeCash
    gSecTypeToString = "Cash"
Case SecTypeCombo
    gSecTypeToString = "Combo"
Case SecTypeIndex
    gSecTypeToString = "Index"
End Select
End Function

Public Function gSecTypeToShortString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    gSecTypeToShortString = "STK"
Case SecTypeFuture
    gSecTypeToShortString = "FUT"
Case SecTypeOption
    gSecTypeToShortString = "OPT"
Case SecTypeFuturesOption
    gSecTypeToShortString = "FOP"
Case SecTypeCash
    gSecTypeToShortString = "CASH"
Case SecTypeCombo
    gSecTypeToShortString = "CMB"
Case SecTypeIndex
    gSecTypeToShortString = "IND"
End Select
End Function

Public Function gUnEscapeRegexSpecialChar( _
                ByRef inString As String) As String
Dim i As Long
Dim s As String
Dim ch As String
Dim skipNextCheck As Boolean

For i = 1 To Len(inString)
    ch = Mid$(inString, i, 1)
    If skipNextCheck Then
        s = s & ch
        skipNextCheck = False
    Else
        If ch = "\" Then
            skipNextCheck = True
        Else
            s = s & ch
        End If
    End If
Next

gUnEscapeRegexSpecialChar = s
End Function

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

gPriceParser.GenerateParsePriceAs32ndsAndFractionsPattern
gPriceParser.GenerateParsePriceAs32ndsPattern
gPriceParser.GenerateParsePriceAs64thsAndFractionsPattern
gPriceParser.GenerateParsePriceAs64thsPattern

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExchangeCode(ByVal code As String)
Const ProcName As String = "addExchangeCode"
Dim failpoint As String
On Error GoTo Err

mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(2 * (UBound(mExchangeCodes) + 1) - 1) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(code)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Private Sub addCurrencyDesc( _
                ByVal code As String, _
                ByVal Description As String)
Const ProcName As String = "addCurrencyDesc"
Dim failpoint As String
On Error GoTo Err

mMaxCurrencyDescsIndex = mMaxCurrencyDescsIndex + 1
If mMaxCurrencyDescsIndex > UBound(mCurrencyDescs) Then
    ReDim Preserve mCurrencyDescs(2 * (UBound(mCurrencyDescs) + 1) - 1) As CurrencyDescriptor
End If
mCurrencyDescs(mMaxCurrencyDescsIndex).code = UCase$(code)
mCurrencyDescs(mMaxCurrencyDescsIndex).Description = UCase$(Description)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Private Function getCurrencyIndex(ByVal code As String) As Long
Dim bottom As Long
Dim top As Long
Dim middle As Long

Const ProcName As String = "getCurrencyIndex"
Dim failpoint As String
On Error GoTo Err

If mMaxCurrencyDescsIndex = 0 Then setupCurrencyDescs

getCurrencyIndex = -1

code = UCase$(code)
bottom = 0
top = mMaxCurrencyDescsIndex
middle = Fix((bottom + top) / 2)

Do
    If code < mCurrencyDescs(middle).code Then
        top = middle
    ElseIf code > mCurrencyDescs(middle).code Then
        bottom = middle
    Else
        getCurrencyIndex = middle
        Exit Function
    End If
    middle = Fix((bottom + top) / 2)
Loop Until bottom = middle

If code = mCurrencyDescs(middle).code Then getCurrencyIndex = middle

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Private Function makeStringArray(ByRef inString) As String()
makeStringArray = Split(inString, ",")
End Function

Private Sub setupExchangeCodes()
Const ProcName As String = "setupExchangeCodes"
Dim failpoint As String
On Error GoTo Err

ReDim mExchangeCodes(31) As String
mMaxExchangeCodesIndex = -1

addExchangeCode "ACE"
addExchangeCode "AEB"
addExchangeCode "AMEX"
addExchangeCode "ARCA"
addExchangeCode "ASX"

addExchangeCode "BATS"
addExchangeCode "BELFOX"
addExchangeCode "BEX"
addExchangeCode "BOX"
addExchangeCode "BRUT"
addExchangeCode "BTRADE"
addExchangeCode "BVME"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CBOT"
addExchangeCode "CBSX"
addExchangeCode "CDE"
addExchangeCode "CFE"
addExchangeCode "CHX"
addExchangeCode "CSFBALGO"

addExchangeCode "DRCTEDGE"
addExchangeCode "DTB"

addExchangeCode "EBS"
addExchangeCode "ECBOT"
addExchangeCode "EDGEA"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FTA"
addExchangeCode "FWB"

addExchangeCode "GLOBEX"

addExchangeCode "HKFE"

addExchangeCode "IBIS"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "JEFFALGO"

addExchangeCode "LAVA"
addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"
addExchangeCode "LSSF"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MEXI"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NQLX"
addExchangeCode "NSE"
addExchangeCode "NSX"
addExchangeCode "NYBOT"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"

addExchangeCode "OMS"
addExchangeCode "ONE"
addExchangeCode "OSE.JPN"

addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TRACKECN"
addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Sub

Private Sub setupCurrencyDescs()
Const ProcName As String = "setupCurrencyDescs"
Dim failpoint As String
On Error GoTo Err

ReDim mCurrencyDescs(127) As CurrencyDescriptor
mMaxCurrencyDescsIndex = -1

addCurrencyDesc "AED", "United Arab Emirates, Dirhams"
addCurrencyDesc "AFN", "Afghanistan, Afghanis"
addCurrencyDesc "ALL", "Albania, Leke"
addCurrencyDesc "AMD", "Armenia, Drams"
addCurrencyDesc "ANG", "Netherlands Antilles, Guilders"
addCurrencyDesc "AOA", "Angola, Kwanza"
addCurrencyDesc "ARS", "Argentina, Pesos"
addCurrencyDesc "AUD", "Australia, Dollars"
addCurrencyDesc "AWG", "Aruba, Guilders"
addCurrencyDesc "AZN", "Azerbaijan, New Manats"
addCurrencyDesc "BAM", "Bosnia and Herzegovina, Convertible Marka"
addCurrencyDesc "BBD", "Barbados, Dollars"
addCurrencyDesc "BDT", "Bangladesh, Taka"
addCurrencyDesc "BGN", "Bulgaria, Leva"
addCurrencyDesc "BHD", "Bahrain, Dinars"
addCurrencyDesc "BIF", "Burundi, Francs"
addCurrencyDesc "BMD", "Bermuda, Dollars"
addCurrencyDesc "BND", "Brunei Darussalam, Dollars"
addCurrencyDesc "BOB", "Bolivia, Bolivianos"
addCurrencyDesc "BRL", "Brazil, Brazil Real"
addCurrencyDesc "BSD", "Bahamas, Dollars"
addCurrencyDesc "BTN", "Bhutan, Ngultrum"
addCurrencyDesc "BWP", "Botswana, Pulas"
addCurrencyDesc "BYR", "Belarus, Rubles"
addCurrencyDesc "BZD", "Belize, Dollars"
addCurrencyDesc "CAD", "Canada, Dollars"
addCurrencyDesc "CDF", "Congo/Kinshasa, Congolese Francs"
addCurrencyDesc "CHF", "Switzerland, Francs"
addCurrencyDesc "CLP", "Chile, Pesos"
addCurrencyDesc "CNY", "China, Yuan Renminbi"
addCurrencyDesc "COP", "Colombia, Pesos"
addCurrencyDesc "CRC", "Costa Rica, Colones"
addCurrencyDesc "CUP", "Cuba, Pesos"
addCurrencyDesc "CVE", "Cape Verde, Escudos"
addCurrencyDesc "CYP", "Cyprus, Pounds"
addCurrencyDesc "CZK", "Czech Republic, Koruny"
addCurrencyDesc "DJF", "Djibouti, Francs"
addCurrencyDesc "DKK", "Denmark, Kroner"
addCurrencyDesc "DOP", "Dominican Republic, Pesos"
addCurrencyDesc "DZD", "Algeria, Algeria Dinars"
addCurrencyDesc "EEK", "Estonia, Krooni"
addCurrencyDesc "EGP", "Egypt, Pounds"
addCurrencyDesc "ERN", "Eritrea, Nakfa"
addCurrencyDesc "ETB", "Ethiopia, Birr"
addCurrencyDesc "EUR", "Euro Member Countries, Euro"
addCurrencyDesc "FJD", "Fiji, Dollars"
addCurrencyDesc "FKP", "Falkland Islands (Malvinas), Pounds"
addCurrencyDesc "GBP", "United Kingdom, Pounds"
addCurrencyDesc "GEL", "Georgia, Lari"
addCurrencyDesc "GGP", "Guernsey, Pounds"
addCurrencyDesc "GHC", "Ghana, Cedis"
addCurrencyDesc "GHS", "Ghana, Cedis"
addCurrencyDesc "GIP", "Gibraltar, Pounds"
addCurrencyDesc "GMD", "Gambia, Dalasi"
addCurrencyDesc "GNF", "Guinea, Francs"
addCurrencyDesc "GTQ", "Guatemala, Quetzales"
addCurrencyDesc "GYD", "Guyana, Dollars"
addCurrencyDesc "HKD", "Hong Kong, Dollars"
addCurrencyDesc "HNL", "Honduras, Lempiras"
addCurrencyDesc "HRK", "Croatia, Kuna"
addCurrencyDesc "HTG", "Haiti, Gourdes"
addCurrencyDesc "HUF", "Hungary, Forint"
addCurrencyDesc "IDR", "Indonesia, Rupiahs"
addCurrencyDesc "ILS", "Israel, New Shekels"
addCurrencyDesc "IMP", "Isle of Man, Pounds"
addCurrencyDesc "INR", "India, Rupees"
addCurrencyDesc "IQD", "Iraq, Dinars"
addCurrencyDesc "IRR", "Iran, Rials"
addCurrencyDesc "ISK", "Iceland, Kronur"
addCurrencyDesc "JEP", "Jersey, Pounds"
addCurrencyDesc "JMD", "Jamaica, Dollars"
addCurrencyDesc "JOD", "Jordan, Dinars"
addCurrencyDesc "JPY", "Japan, Yen"
addCurrencyDesc "KES", "Kenya, Shillings"
addCurrencyDesc "KGS", "Kyrgyzstan, Soms"
addCurrencyDesc "KHR", "Cambodia, Riels"
addCurrencyDesc "KMF", "Comoros, Francs"
addCurrencyDesc "KPW", "Korea (North), Won"
addCurrencyDesc "KRW", "Korea (South), Won"
addCurrencyDesc "KWD", "Kuwait, Dinars"
addCurrencyDesc "KYD", "Cayman Islands, Dollars"
addCurrencyDesc "KZT", "Kazakhstan, Tenge"
addCurrencyDesc "LAK", "Laos, Kips"
addCurrencyDesc "LBP", "Lebanon, Pounds"
addCurrencyDesc "LKR", "Sri Lanka, Rupees"
addCurrencyDesc "LRD", "Liberia, Dollars"
addCurrencyDesc "LSL", "Lesotho, Maloti"
addCurrencyDesc "LTL", "Lithuania, Litai"
addCurrencyDesc "LVL", "Latvia, Lati"
addCurrencyDesc "LYD", "Libya, Dinars"
addCurrencyDesc "MAD", "Morocco, Dirhams"
addCurrencyDesc "MDL", "Moldova, Lei"
addCurrencyDesc "MGA", "Madagascar, Ariary"
addCurrencyDesc "MKD", "Macedonia, Denars"
addCurrencyDesc "MMK", "Myanmar (Burma), Kyats"
addCurrencyDesc "MNT", "Mongolia, Tugriks"
addCurrencyDesc "MOP", "Macau, Patacas"
addCurrencyDesc "MRO", "Mauritania, Ouguiyas"
addCurrencyDesc "MTL", "Malta, Liri"
addCurrencyDesc "MUR", "Mauritius, Rupees"
addCurrencyDesc "MVR", "Maldives (Maldive Islands), Rufiyaa"
addCurrencyDesc "MWK", "Malawi, Kwachas"
addCurrencyDesc "MXN", "Mexico, Pesos"
addCurrencyDesc "MYR", "Malaysia, Ringgits"
addCurrencyDesc "MZN", "Mozambique, Meticais"
addCurrencyDesc "NAD", "Namibia, Dollars"
addCurrencyDesc "NGN", "Nigeria, Nairas"
addCurrencyDesc "NIO", "Nicaragua, Cordobas"
addCurrencyDesc "NOK", "Norway, Krone"
addCurrencyDesc "NPR", "Nepal, Nepal Rupees"
addCurrencyDesc "NZD", "New Zealand, Dollars"
addCurrencyDesc "OMR", "Oman, Rials"
addCurrencyDesc "PAB", "Panama, Balboa"
addCurrencyDesc "PEN", "Peru, Nuevos Soles"
addCurrencyDesc "PGK", "Papua New Guinea, Kina"
addCurrencyDesc "PHP", "Philippines, Pesos"
addCurrencyDesc "PKR", "Pakistan, Rupees"
addCurrencyDesc "PLN", "Poland, Zlotych"
addCurrencyDesc "PYG", "Paraguay, Guarani"
addCurrencyDesc "QAR", "Qatar, Rials"
addCurrencyDesc "RON", "Romania, New Lei"
addCurrencyDesc "RSD", "Serbia, Dinars"
addCurrencyDesc "RUB", "Russia, Rubles"
addCurrencyDesc "RWF", "Rwanda, Rwanda Francs"
addCurrencyDesc "SAR", "Saudi Arabia, Riyals"
addCurrencyDesc "SBD", "Solomon Islands, Dollars"
addCurrencyDesc "SCR", "Seychelles, Rupees"
addCurrencyDesc "SDG", "Sudan, Pounds"
addCurrencyDesc "SEK", "Sweden, Kronor"
addCurrencyDesc "SGD", "Singapore, Dollars"
addCurrencyDesc "SHP", "Saint Helena, Pounds"
addCurrencyDesc "SKK", "Slovakia, Koruny"
addCurrencyDesc "SLL", "Sierra Leone, Leones"
addCurrencyDesc "SOS", "Somalia, Shillings"
addCurrencyDesc "SPL", "Seborga, Luigini"
addCurrencyDesc "SRD", "Suriname, Dollars"
addCurrencyDesc "STD", "São Tome and Principe, Dobras"
addCurrencyDesc "SVC", "El Salvador, Colones"
addCurrencyDesc "SYP", "Syria, Pounds"
addCurrencyDesc "SZL", "Swaziland, Emalangeni"
addCurrencyDesc "THB", "Thailand, Baht"
addCurrencyDesc "TJS", "Tajikistan, Somoni"
addCurrencyDesc "TMM", "Turkmenistan, Manats"
addCurrencyDesc "TND", "Tunisia, Dinars"
addCurrencyDesc "TOP", "Tonga, Pa'anga"
addCurrencyDesc "TRY", "Turkey, New Lira"
addCurrencyDesc "TTD", "Trinidad and Tobago, Dollars"
addCurrencyDesc "TVD", "Tuvalu, Tuvalu Dollars"
addCurrencyDesc "TWD", "Taiwan, New Dollars"
addCurrencyDesc "TZS", "Tanzania, Shillings"
addCurrencyDesc "UAH", "Ukraine, Hryvnia"
addCurrencyDesc "UGX", "Uganda, Shillings"
addCurrencyDesc "USD", "United States of America, Dollars"
addCurrencyDesc "UYU", "Uruguay, Pesos"
addCurrencyDesc "UZS", "Uzbekistan, Sums"
addCurrencyDesc "VEB", "Venezuela, Bolivares"
addCurrencyDesc "VND", "Viet Nam, Dong"
addCurrencyDesc "VUV", "Vanuatu, Vatu"
addCurrencyDesc "WST", "Samoa, Tala"
addCurrencyDesc "XAF", "Communauté Financière Africaine BEAC, Francs"
addCurrencyDesc "XAG", "Silver, Ounces"
addCurrencyDesc "XAU", "Gold, Ounces"
addCurrencyDesc "XCD", "East Caribbean Dollars"
addCurrencyDesc "XOF", "Communauté Financière Africaine BCEAO, Francs"
addCurrencyDesc "XPD", "Palladium Ounces"
addCurrencyDesc "XPF", "Comptoirs Français du Pacifique Francs"
addCurrencyDesc "XPT", "Platinum, Ounces"
addCurrencyDesc "YER", "Yemen, Rials"
addCurrencyDesc "ZAR", "South Africa, Rand"
addCurrencyDesc "ZMK", "Zambia, Kwacha"
addCurrencyDesc "ZWD", "Zimbabwe, Zimbabwe Dollars"


ReDim Preserve mCurrencyDescs(mMaxCurrencyDescsIndex) As CurrencyDescriptor

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint

End Sub
