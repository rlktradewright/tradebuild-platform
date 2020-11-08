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

Public Const ProjectName                                As String = "ContractUtils27"
Private Const ModuleName                                As String = "Globals"

Public Const ConfigSectionContractSpecifier             As String = "Specifier"

Public Const ConfigSettingContractSpecCurrency          As String = "&Currency"
Public Const ConfigSettingContractSpecExpiry            As String = "&Expiry"
Public Const ConfigSettingContractSpecExchange          As String = "&Exchange"
Public Const ConfigSettingContractSpecLocalSymbol       As String = "&LocalSymbol"
Public Const ConfigSettingContractSpecMultiplier        As String = "&Multiplier"
Public Const ConfigSettingContractSpecRight             As String = "&Right"
Public Const ConfigSettingContractSpecSecType           As String = "&SecType"
Public Const ConfigSettingContractSpecStrikePrice       As String = "&StrikePrice"
Public Const ConfigSettingContractSpecSymbol            As String = "&Symbol"

Public Const ConfigSettingDaysBeforeExpiryToSwitch      As String = "&DaysBeforeExpiryToSwitch"
Public Const ConfigSettingDescription                   As String = "&Description"
Public Const ConfigSettingExpiryDate                    As String = "&ExpiryDate"
Public Const ConfigSettingFullSessionEndTime            As String = "&FullSessionEndTime"
Public Const ConfigSettingFullSessionStartTime          As String = "&FullSessionStartTime"
Public Const ConfigSettingMultiplier                    As String = "&Multiplier"
Public Const ConfigSettingSessionEndTime                As String = "&SessionEndTime"
Public Const ConfigSettingSessionStartTime              As String = "&SessionStartTime"
Public Const ConfigSettingTickSize                      As String = "&TickSize"
Public Const ConfigSettingTimezoneName                  As String = "&Timezone"

Public Const DefaultDaysBeforeExpiryToSwitch            As Long = 0
Public Const DefaultExpiry                              As String = "1899-12-30"
Public Const DefaultMultiplier                          As Double = 1#
Public Const DefaultTickSize                            As Double = 0.01
Public Const DefaultTimezoneName                        As String = "Eastern Standard Time"

Public Const MaxContractExpiryOffset                    As Long = 10
Public Const MaxContractDaysBeforeExpiryToSwitch        As Long = 20

'@================================================================================
' Member variables
'@================================================================================

Private mExchangeCodesInitialised                       As Boolean
Private mExchangeCodes()                                As String
Private mMaxExchangeCodesIndex                          As Long

Private mCurrencyDescsColl                              As SortedDictionary
Private mCurrencyDescriptors()                          As CurrencyDescriptor

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

Public Function gContractSpecsCompatible( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "gContractSpecsCompatible"
On Error GoTo Err

If pContractSpec1.LocalSymbol <> "" And UCase$(pContractSpec1.LocalSymbol) <> UCase$(pContractSpec2.LocalSymbol) Then Exit Function
If pContractSpec1.Symbol <> "" And UCase$(pContractSpec1.Symbol) <> UCase$(pContractSpec2.Symbol) Then Exit Function
If pContractSpec1.SecType <> SecTypeNone And pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
If pContractSpec1.Exchange <> "" And UCase$(pContractSpec1.Exchange) <> UCase$(pContractSpec2.Exchange) Then Exit Function
If pContractSpec1.Expiry <> "" And pContractSpec1.Expiry <> Left$(pContractSpec2.Expiry, Len(pContractSpec1.Expiry)) Then Exit Function
If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
If pContractSpec1.CurrencyCode <> "" And UCase$(pContractSpec1.CurrencyCode) <> UCase$(pContractSpec2.CurrencyCode) Then Exit Function
If pContractSpec1.Right <> OptNone And pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
If pContractSpec1.Strike <> 0# And pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function

gContractSpecsCompatible = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractSpecsCompare( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier, _
                ByRef pSortkeys() As ContractSortKeyIds, _
                ByVal pAscending As Boolean) As Long
Const ProcName As String = "gContractSpecsCompare"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(pSortkeys)
    Select Case pSortkeys(i)
    Case ContractSortKeyNone
        Exit Function
    Case ContractSortKeyLocalSymbol
        gContractSpecsCompare = StrComp(pContractSpec1.LocalSymbol, pContractSpec2.LocalSymbol, vbTextCompare)
    Case ContractSortKeySymbol
        gContractSpecsCompare = StrComp(pContractSpec1.Symbol, pContractSpec2.Symbol, vbTextCompare)
    Case ContractSortKeySecType
        gContractSpecsCompare = StrComp(gSecTypeToShortString(pContractSpec1.SecType), gSecTypeToShortString(pContractSpec2.SecType), vbTextCompare)
    Case ContractSortKeyExchange
        gContractSpecsCompare = StrComp(pContractSpec1.Exchange, pContractSpec2.Exchange, vbTextCompare)
    Case ContractSortKeyExpiry
        gContractSpecsCompare = StrComp(pContractSpec1.Expiry, pContractSpec2.Expiry, vbTextCompare)
    Case ContractSortKeyMultiplier
        gContractSpecsCompare = StrComp(pContractSpec1.Multiplier, pContractSpec2.Multiplier, vbBinaryCompare)
    Case ContractSortKeyCurrency
        gContractSpecsCompare = StrComp(pContractSpec1.CurrencyCode, pContractSpec2.CurrencyCode, vbTextCompare)
    Case ContractSortKeyRight
        gContractSpecsCompare = StrComp(gOptionRightToString(pContractSpec1.Right), gOptionRightToString(pContractSpec2.Right), vbTextCompare)
    Case ContractSortKeyStrike
        gContractSpecsCompare = Sgn(pContractSpec1.Strike - pContractSpec2.Strike)
    End Select
    If gContractSpecsCompare <> 0 Then
        If Not pAscending Then gContractSpecsCompare = -gContractSpecsCompare
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractSpecsEqual( _
                ByVal pContractSpec1 As IContractSpecifier, _
                ByVal pContractSpec2 As IContractSpecifier) As Boolean
Const ProcName As String = "gContractSpecsEqual"
On Error GoTo Err

If pContractSpec1 Is Nothing Then Exit Function
If pContractSpec2 Is Nothing Then Exit Function
If pContractSpec1 Is pContractSpec2 Then
    gContractSpecsEqual = True
Else
    gContractSpecsEqual = (gGetContractSpecKey(pContractSpec1) = gGetContractSpecKey(pContractSpec2))
    If pContractSpec1.CurrencyCode <> pContractSpec2.CurrencyCode Then Exit Function
    If pContractSpec1.Exchange <> pContractSpec2.Exchange Then Exit Function
    If pContractSpec1.Expiry <> pContractSpec2.Expiry Then Exit Function
    If pContractSpec1.LocalSymbol <> pContractSpec2.LocalSymbol Then Exit Function
    If pContractSpec1.Multiplier <> pContractSpec2.Multiplier Then Exit Function
    If pContractSpec1.Right <> pContractSpec2.Right Then Exit Function
    If pContractSpec1.SecType <> pContractSpec2.SecType Then Exit Function
    If pContractSpec1.Strike <> pContractSpec2.Strike Then Exit Function
    If pContractSpec1.Symbol <> pContractSpec2.Symbol Then Exit Function
    gContractSpecsEqual = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractToString(ByVal pContract As IContract) As String
Const ProcName As String = "gContractToString"
On Error GoTo Err

gContractToString = "Specifier=(" & pContract.Specifier.ToString & "); " & _
            "Description=" & pContract.Description & "; " & _
            "Expiry date=" & pContract.ExpiryDate & "; " & _
            "Tick size=" & pContract.TickSize & "; " & _
            "Session start=" & FormatDateTime(pContract.SessionStartTime, vbShortTime) & "; " & _
            "Session end=" & FormatDateTime(pContract.SessionEndTime, vbShortTime) & "; " & _
            "Full session start=" & FormatDateTime(pContract.FullSessionStartTime, vbShortTime) & "; " & _
            "Full session end=" & FormatDateTime(pContract.FullSessionEndTime, vbShortTime) & "; " & _
            "TimezoneName=" & pContract.TimezoneName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gContractToXML(ByVal pContract As IContract) As String
Const ProcName As String = "gContractToXML"
On Error GoTo Err

Dim XMLdoc As DOMDocument30: Set XMLdoc = New DOMDocument30
Dim lContractElement As IXMLDOMElement: Set lContractElement = XMLdoc.createElement("contract")
Set XMLdoc.documentElement = lContractElement
lContractElement.setAttribute "xmlns", "urn:tradewright.com:tradebuild"
lContractElement.setAttribute "minimumtick", pContract.TickSize
lContractElement.setAttribute "sessionstarttime", Format(pContract.SessionStartTime, "hh:mm:ss")
lContractElement.setAttribute "sessionendtime", Format(pContract.SessionEndTime, "hh:mm:ss")
lContractElement.setAttribute "fullsessionstarttime", Format(pContract.FullSessionStartTime, "hh:mm:ss")
lContractElement.setAttribute "fullsessionendtime", Format(pContract.FullSessionEndTime, "hh:mm:ss")
lContractElement.setAttribute "Description", pContract.Description
lContractElement.setAttribute "numberofdecimals", pContract.NumberOfDecimals
lContractElement.setAttribute "timezonename", pContract.TimezoneName

Dim Specifier As IXMLDOMElement: Set Specifier = XMLdoc.createElement("specifier")
lContractElement.appendChild Specifier
Specifier.setAttribute "symbol", pContract.Specifier.Symbol
Specifier.setAttribute "sectype", pContract.Specifier.SecType
Specifier.setAttribute "expiry", pContract.Specifier.Expiry
Specifier.setAttribute "exchange", pContract.Specifier.Exchange
Specifier.setAttribute "currencycode", pContract.Specifier.CurrencyCode
Specifier.setAttribute "localsymbol", pContract.Specifier.LocalSymbol
Specifier.setAttribute "multiplier", pContract.Specifier.Multiplier
Specifier.setAttribute "right", pContract.Specifier.Right
Specifier.setAttribute "strike", pContract.Specifier.Strike

Dim ComboLegs As IXMLDOMElement: Set ComboLegs = XMLdoc.createElement("combolegs")
Specifier.appendChild ComboLegs
If Not pContract.Specifier.ComboLegs Is Nothing Then
    Dim comboLegObj As ComboLeg
    For Each comboLegObj In pContract.Specifier.ComboLegs
        Dim ComboLeg As IXMLDOMElement: Set ComboLeg = XMLdoc.createElement("comboleg")
        ComboLegs.appendChild ComboLeg
        
        Set Specifier = XMLdoc.createElement("specifier")
        ComboLeg.appendChild Specifier
        ComboLeg.setAttribute "isBuyLeg", comboLegObj.IsBuyLeg
        ComboLeg.setAttribute "Ratio", comboLegObj.Ratio
        Specifier.setAttribute "symbol", pContract.Specifier.Symbol
        Specifier.setAttribute "sectype", pContract.Specifier.SecType
        Specifier.setAttribute "expiry", pContract.Specifier.Expiry
        Specifier.setAttribute "exchange", pContract.Specifier.Exchange
        Specifier.setAttribute "currencycode", pContract.Specifier.CurrencyCode
        Specifier.setAttribute "localsymbol", pContract.Specifier.LocalSymbol
        Specifier.setAttribute "right", pContract.Specifier.Right
        Specifier.setAttribute "strike", pContract.Specifier.Strike
    Next
End If
gContractToXML = XMLdoc.xml

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateContractSpecifier( _
                ByVal LocalSymbol As String, _
                ByVal Symbol As String, _
                ByVal Exchange As String, _
                ByVal SecType As SecurityTypes, _
                ByVal CurrencyCode As String, _
                ByVal Expiry As String, _
                ByVal Multiplier As Double, _
                ByVal Strike As Double, _
                ByVal Right As OptionRights) As IContractSpecifier
Const ProcName As String = "gCreateContractSpecifier"
On Error GoTo Err

Dim lSpec As New ContractSpecifier
lSpec.Initialise LocalSymbol, Symbol, Exchange, SecType, CurrencyCode, Expiry, Multiplier, Strike, Right
Set gCreateContractSpecifier = lSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateContractSpecifierFromString(ByVal pSpecString As String) As IContractSpecifier
Const ProcName As String = "gCreateContractSpecifierFromString"
On Error GoTo Err

Const ContractSpecRegex As String = "^(STK|FUT|OPT|FOP|CASH|COMBO|INDEX)\:([a-zA-Z0-9][ a-zA-Z0-9\-]+)(?:@([a-zA-Z]+))?(?:\(([a-zA-Z]+)\))?$"
gRegExp.Pattern = ContractSpecRegex
gRegExp.IgnoreCase = True

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(pSpecString)

AssertArgument lMatches.Count = 1, "The contract must be specified: <sectype>:<localsymbol>[@<exchange>][(<currency>)], eg STK:MSFT@SMART(USD)"

Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lSecType As SecurityTypes
lSecType = gSecTypeFromString(lMatch.SubMatches(0))
AssertArgument lSecType <> SecTypeNone, "A valid security type must be supplied"

Dim lLocalSymbol As String
lLocalSymbol = lMatch.SubMatches(1)

Dim lExchange As String
lExchange = lMatch.SubMatches(2)

Dim lCurrency As String
lCurrency = lMatch.SubMatches(3)

Set gCreateContractSpecifierFromString = gCreateContractSpecifier( _
                                                lLocalSymbol, _
                                                "", _
                                                lExchange, _
                                                lSecType, _
                                                lCurrency, _
                                                "", _
                                                1#, _
                                                0#, _
                                                OptNone)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetContractSpecKey(ByVal pSpec As IContractSpecifier) As String
Const ProcName As String = "gGetContractSpecKey"
On Error GoTo Err

gGetContractSpecKey = pSpec.LocalSymbol & "|" & _
    CStr(pSpec.SecType) & "|" & _
    pSpec.Symbol & "|" & _
    pSpec.Expiry & "|" & _
    pSpec.Strike & "|" & _
    CStr(pSpec.Right) & "|" & _
    pSpec.Exchange & "|" & _
    pSpec.CurrencyCode & "|" & _
    pSpec.Multiplier

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetCurrencyDescriptor( _
                ByVal Code As String) As CurrencyDescriptor
Const ProcName As String = "gGetCurrencyDescriptor"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then setupCurrencyDescs

Code = UCase$(Code)
AssertArgument mCurrencyDescsColl.Contains(Code), "Invalid currency Code"

gGetCurrencyDescriptor = mCurrencyDescsColl.Item(Code)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetCurrencyDescriptors() As CurrencyDescriptor()
Const ProcName As String = "gGetCurrencyDescriptors"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then
    setupCurrencyDescs
    ReDim mCurrencyDescriptors(mCurrencyDescsColl.Count - 1) As CurrencyDescriptor
    Dim lDesc As Variant
    Dim i As Long
    For Each lDesc In mCurrencyDescsColl
        mCurrencyDescriptors(i) = lDesc
        i = i + 1
    Next
End If
gGetCurrencyDescriptors = mCurrencyDescriptors

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetExchangeCodes() As String()
Const ProcName As String = "gGetExchangeCodes"
On Error GoTo Err

If Not mExchangeCodesInitialised Then setupExchangeCodes
gGetExchangeCodes = mExchangeCodes

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Function gLoadContractFromConfig(ByVal pConfig As ConfigurationSection) As IContract
Const ProcName As String = "gLoadContractFromConfig"
On Error GoTo Err

Dim lSpec As ContractSpecifier
Set lSpec = gLoadContractSpecFromConfig(pConfig.AddConfigurationSection(ConfigSectionContractSpecifier))

Dim lContract As New Contract
lContract.Specifier = lSpec

With pConfig
    lContract.DaysBeforeExpiryToSwitch = .GetSetting(ConfigSettingDaysBeforeExpiryToSwitch, DefaultDaysBeforeExpiryToSwitch)
    lContract.Description = .GetSetting(ConfigSettingDescription, "")
    lContract.ExpiryDate = CDate(.GetSetting(ConfigSettingExpiryDate, DefaultExpiry))
    lContract.FullSessionEndTime = .GetSetting(ConfigSettingFullSessionEndTime, "00:00:00")
    lContract.FullSessionStartTime = .GetSetting(ConfigSettingFullSessionStartTime, "00:00:00")
    lContract.SessionEndTime = .GetSetting(ConfigSettingSessionEndTime, "00:00:00")
    lContract.SessionStartTime = .GetSetting(ConfigSettingSessionStartTime, "00:00:00")
    lContract.TickSize = .GetSetting(ConfigSettingTickSize, DefaultTickSize)
    lContract.TimezoneName = .GetSetting(ConfigSettingTimezoneName, DefaultTimezoneName)
    If .GetSetting(ConfigSettingMultiplier, DefaultMultiplier) <> DefaultMultiplier Then
        lSpec.Multiplier = .GetSetting(ConfigSettingMultiplier, DefaultMultiplier)
    End If
End With

Set gLoadContractFromConfig = lContract

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadContractSpecFromConfig(ByVal pConfig As ConfigurationSection) As ContractSpecifier
Const ProcName As String = "gLoadContractSpecFromConfig"
On Error GoTo Err

Dim lContractSpec As ContractSpecifier
With pConfig
    Set lContractSpec = gCreateContractSpecifier(.GetSetting(ConfigSettingContractSpecLocalSymbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecSymbol, ""), _
                                            .GetSetting(ConfigSettingContractSpecExchange, ""), _
                                            gSecTypeFromString(.GetSetting(ConfigSettingContractSpecSecType, "")), _
                                            .GetSetting(ConfigSettingContractSpecCurrency, ""), _
                                            .GetSetting(ConfigSettingContractSpecExpiry, ""), _
                                            CDbl(.GetSetting(ConfigSettingContractSpecMultiplier, DefaultMultiplier)), _
                                            CDbl(.GetSetting(ConfigSettingContractSpecStrikePrice, "0.0")), _
                                            gOptionRightFromString(.GetSetting(ConfigSettingContractSpecRight, "")))

End With

Set gLoadContractSpecFromConfig = lContractSpec

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gIsContractExpired(ByVal pContract As IContract) As Boolean
Const ProcName As String = "gIsContractExpired"
On Error GoTo Err

Select Case pContract.Specifier.SecType
Case SecTypeFuture, SecTypeOption, SecTypeFuturesOption
    If Int(pContract.ExpiryDate) < Int(Now) Then gIsContractExpired = True
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsValidCurrencyCode(ByVal Code As String) As Boolean
Const ProcName As String = "gIsValidCurrencyCode"
On Error GoTo Err

If mCurrencyDescsColl Is Nothing Then setupCurrencyDescs

gIsValidCurrencyCode = mCurrencyDescsColl.Contains(UCase$(Code))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsValidExchangeCode(ByVal Code As String) As Boolean
Const ProcName As String = "gIsValidExchangeCode"
On Error GoTo Err

If Not mExchangeCodesInitialised Then setupExchangeCodes

Code = UCase$(Code)

gIsValidExchangeCode = BinarySearchStrings( _
                            Code, _
                            mExchangeCodes, _
                            IsCaseSensitive:=False) >= 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsValidExchangeSpecifier(ByVal pExchangeSpec As String) As Boolean
Const ProcName As String = "gIsValidExchangeSpecifier"
On Error GoTo Err

pExchangeSpec = UCase$(pExchangeSpec)

Const ExchangeSmartQualified As String = "SMART/"
If InStr(1, pExchangeSpec, ExchangeSmartQualified) <> 1 Then
    gIsValidExchangeSpecifier = gIsValidExchangeCode(pExchangeSpec)
    Exit Function
End If

gIsValidExchangeSpecifier = gIsValidExchangeCode( _
                                Right$(pExchangeSpec, _
                                        Len(pExchangeSpec) - Len(ExchangeSmartQualified)))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsOffsetExpiry( _
                ByVal Value As String) As Boolean
Dim l1 As Long
Dim l2 As Long
gIsOffsetExpiry = gParseOffsetExpiry(Value, l1, l2)
End Function

Public Function gIsValidExpiry( _
                ByVal Value As String) As Boolean
Const ProcName As String = "gIsValidExpiry"
On Error GoTo Err

If gIsOffsetExpiry(Value) Then
    gIsValidExpiry = True
    Exit Function
End If

Dim d As Date

If IsDate(Value) Then
    d = CDate(Value)
ElseIf Len(Value) = 8 Then
    Dim datestring As String
    datestring = Left$(Value, 4) & "/" & Mid$(Value, 5, 2) & "/" & Right$(Value, 2)
    If IsDate(datestring) Then d = CDate(datestring)
End If

If d <> 0 Then
    If d >= CDate((Year(Now) - 20) & "/01/01") And d <= CDate((Year(Now) + 10) & "/12/31") Then
        gIsValidExpiry = True
        Exit Function
    End If
End If

If Len(Value) = 6 Then
    If IsInteger(Value, (Year(Now) - 20) * 100 + 1, (Year(Now) + 10) * 100 + 12) Then
        If Right$(Value, 2) <= 12 Then
            gIsValidExpiry = True
            Exit Function
        End If
    End If
End If

gIsValidExpiry = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsValidPrice( _
                ByVal pPrice As Double, _
                ByVal pPrevValidPrice As Double, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double) As Boolean
Const ProcName As String = "gIsValidPrice"
On Error GoTo Err

If pTickSize = 0 Then
    gIsValidPrice = True
    Exit Function
End If

If pPrevValidPrice = 0 Or pPrevValidPrice = MaxDouble Then
    If Abs(pPrice) / pTickSize > &H3FFFFF Then Exit Function ' note that Z index has ticksize 0.01
                                                            ' so we need to allow plenty of room
                                                            ' &H3FFFFF = 4194303
    
    ' A first price of 0 is always considered invalid. Although some indexes can validly
    ' have zero values, it is unlikely on the very first price notified, and this check
    ' catches the occasional zero prices sent by IB when, for example, the Z ticker
    ' is started
    If pPrice = 0 Then Exit Function
    
    If pSecType = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive, zero and negative values
    Else
        If pPrice < 0 Then Exit Function
        If pPrice / pTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
    End If
Else
    'If Abs(pPrevValidPrice - pPrice) / pTickSize > 32767 Then Exit Function
    
    If pSecType = SecTypeIndex Then
        ' don't do this check for indexes because some of them, such as TICK-NYSE, can have both
        ' positive and negative values - moreover the value can change dramatically from one
        ' tick to the next
    Else
        If pPrice <= 0 Then Exit Function
        If pPrice / pTickSize < 1 Then Exit Function ' catch occasional very small prices from IB
        'If pPrice < (2 * pPrevValidPrice) / 3 Or pPrice > (3 * pPrevValidPrice) / 2 Then Exit Function
    End If
    
End If

gIsValidPrice = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsValidSecType( _
                ByVal Value As Long) As Boolean
gIsValidSecType = True
Select Case Value
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

Public Function gOptionRightFromString(ByVal Value As String) As OptionRights
Select Case UCase$(Value)
Case ""
    gOptionRightFromString = OptNone
Case "CALL", "C"
    gOptionRightFromString = OptCall
Case "PUT", "P"
    gOptionRightFromString = OptPut
End Select
End Function

Public Function gOptionRightToString(ByVal Value As OptionRights) As String
Select Case Value
Case OptNone
    gOptionRightToString = ""
Case OptCall
    gOptionRightToString = "Call"
Case OptPut
    gOptionRightToString = "Put"
End Select
End Function

Public Function gParseOffsetExpiry( _
                ByVal Value As String, _
                ByRef pExpiryOffset As Long, _
                ByRef pDaysBeforeExpiryToSwitch As Long) As Boolean
Const OffsetExpiryFormat As String = "^(\d\d?)(?:\[(\d\d?)d\])?$"

gRegExp.Pattern = OffsetExpiryFormat
gRegExp.IgnoreCase = True

Dim lMatches As MatchCollection
Set lMatches = gRegExp.Execute(Trim$(Value))

If lMatches.Count <> 1 Then Exit Function

Dim lResult As Boolean: lResult = True
Dim lMatch As Match: Set lMatch = lMatches(0)

Dim lOffsetStr As String
lOffsetStr = lMatch.SubMatches(0)
If lOffsetStr = "" Then
    pExpiryOffset = 0
ElseIf IsInteger(lOffsetStr, 0, MaxContractExpiryOffset) Then
    pExpiryOffset = CInt(lOffsetStr)
Else
    lResult = False
End If

Dim lDaysBeforeExpiryStr As String
lDaysBeforeExpiryStr = lMatch.SubMatches(1)
If lDaysBeforeExpiryStr = "" Then
    pDaysBeforeExpiryToSwitch = 0
ElseIf IsInteger(lDaysBeforeExpiryStr, 0, MaxContractDaysBeforeExpiryToSwitch) Then
    pDaysBeforeExpiryToSwitch = CInt(lDaysBeforeExpiryStr)
Else
    lResult = False
End If

gParseOffsetExpiry = lResult
End Function

Public Property Get gRegExp() As RegExp
Const ProcName As String = "gRegExp"
On Error GoTo Err

Static sRegExp As RegExp
If sRegExp Is Nothing Then Set sRegExp = New RegExp
Set gRegExp = sRegExp

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Sub gSaveContractSpecToConfig(ByVal pContractSpec As IContractSpecifier, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "gSaveContractSpecToConfig"
On Error GoTo Err

With pConfig
    .SetSetting ConfigSettingContractSpecLocalSymbol, pContractSpec.LocalSymbol
    .SetSetting ConfigSettingContractSpecSymbol, pContractSpec.Symbol
    .SetSetting ConfigSettingContractSpecExchange, pContractSpec.Exchange
    .SetSetting ConfigSettingContractSpecSecType, gSecTypeToString(pContractSpec.SecType)
    .SetSetting ConfigSettingContractSpecCurrency, pContractSpec.CurrencyCode
    .SetSetting ConfigSettingContractSpecExpiry, pContractSpec.Expiry
    .SetSetting ConfigSettingContractSpecMultiplier, pContractSpec.Multiplier
    .SetSetting ConfigSettingContractSpecStrikePrice, pContractSpec.Strike
    .SetSetting ConfigSettingContractSpecRight, gOptionRightToString(pContractSpec.Right)
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSaveContractToConfig(ByVal pContract As IContract, ByVal pConfig As ConfigurationSection)
Const ProcName As String = "gSaveContractToConfig"
On Error GoTo Err

gSaveContractSpecToConfig pContract.Specifier, pConfig.AddConfigurationSection(ConfigSectionContractSpecifier)

With pConfig
    .SetSetting ConfigSettingDaysBeforeExpiryToSwitch, pContract.DaysBeforeExpiryToSwitch
    .SetSetting ConfigSettingDescription, pContract.Description
    .SetSetting ConfigSettingExpiryDate, FormatTimestamp(pContract.ExpiryDate, TimestampDateOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionEndTime, FormatTimestamp(pContract.SessionEndTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingSessionStartTime, FormatTimestamp(pContract.SessionStartTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingFullSessionEndTime, FormatTimestamp(pContract.FullSessionEndTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingFullSessionStartTime, FormatTimestamp(pContract.FullSessionStartTime, TimestampTimeOnlyISO8601 + TimestampNoMillisecs)
    .SetSetting ConfigSettingTickSize, pContract.TickSize
    .SetSetting ConfigSettingTimezoneName, pContract.TimezoneName
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gSecTypeFromString(ByVal Value As String) As SecurityTypes
Select Case UCase$(Value)
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
Case Else
    gSecTypeFromString = SecTypeNone
End Select
End Function

Public Function gSecTypeToString(ByVal Value As SecurityTypes) As String
Select Case Value
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

Public Function gSecTypeToShortString(ByVal Value As SecurityTypes) As String
Select Case Value
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

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gUnEscapeRegexSpecialChar( _
                ByRef inString As String) As String
Dim s As String
Dim skipNextCheck As Boolean
Dim i As Long
For i = 1 To Len(inString)
    Dim ch As String: ch = Mid$(inString, i, 1)
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
GPriceParser.gInit
GPriceFormatter.gInit
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addExchangeCode(ByVal Code As String)
Const ProcName As String = "addExchangeCode"
On Error GoTo Err

mMaxExchangeCodesIndex = mMaxExchangeCodesIndex + 1
If mMaxExchangeCodesIndex > UBound(mExchangeCodes) Then
    ReDim Preserve mExchangeCodes(2 * (UBound(mExchangeCodes) + 1) - 1) As String
End If
mExchangeCodes(mMaxExchangeCodesIndex) = UCase$(Code)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addCurrencyDesc( _
                ByVal Code As String, _
                ByVal Description As String)
Const ProcName As String = "addCurrencyDesc"
On Error GoTo Err

Dim lDescriptor As CurrencyDescriptor
Code = UCase$(Code)
lDescriptor.Code = Code
lDescriptor.Description = Description
mCurrencyDescsColl.Add lDescriptor, Code

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupExchangeCodes()
Const ProcName As String = "setupExchangeCodes"
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
addExchangeCode "BYX"

addExchangeCode "CAES"
addExchangeCode "CBOE"
addExchangeCode "CBOE2"
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
addExchangeCode "EDGX"
addExchangeCode "EMERALD"
addExchangeCode "EUREX"
addExchangeCode "EUREXUS"

addExchangeCode "FOXRIVER"
addExchangeCode "FTA"
addExchangeCode "FWB"

addExchangeCode "GEMINI"
addExchangeCode "GLOBEX"

addExchangeCode "HKFE"

addExchangeCode "IBIS"
addExchangeCode "ICEEU"
addExchangeCode "IDEAL"
addExchangeCode "IDEALPRO"
addExchangeCode "IDEM"
addExchangeCode "IEX"
addExchangeCode "INET"
addExchangeCode "INSTINET"
addExchangeCode "ISE"
addExchangeCode "ISLAND"

addExchangeCode "JEFFALGO"

addExchangeCode "LAVA"
addExchangeCode "LIFFE"
addExchangeCode "LIFFE_NF"
addExchangeCode "LSE"
addExchangeCode "LSEIOB1"
addExchangeCode "LSSF"

addExchangeCode "MATIF"
addExchangeCode "MEFF"
addExchangeCode "MEFFRV"
addExchangeCode "MERCURY"
addExchangeCode "MEXI"
addExchangeCode "MIAX"
addExchangeCode "MONEP"
addExchangeCode "MXT"

addExchangeCode "NASDAQ"
addExchangeCode "NASDAQBX"
addExchangeCode "NASDAQOM"
addExchangeCode "NQLX"
addExchangeCode "NSE"
addExchangeCode "NSX"
addExchangeCode "NYBOT"
addExchangeCode "NYMEX"
addExchangeCode "NYSE"
addExchangeCode "NYSENAT"

addExchangeCode "OMS"
addExchangeCode "OMXNO"
addExchangeCode "ONE"
addExchangeCode "OSE"
addExchangeCode "OSE.JPN"

addExchangeCode "PEARL"
addExchangeCode "PHLX"
addExchangeCode "PINK"
addExchangeCode "PSE"
addExchangeCode "PSX"
addExchangeCode "PURE"

addExchangeCode "QBALGO"

addExchangeCode "RDBK"

addExchangeCode "SBF"
addExchangeCode "SEHK"
addExchangeCode "SFB"
addExchangeCode "SGX"
addExchangeCode "SMART"
addExchangeCode "SMARTCAN"
addExchangeCode "SMARTEUR"
addExchangeCode "SMARTNASDAQ"
addExchangeCode "SMARTNYSE"
addExchangeCode "SMARTUK"
addExchangeCode "SMARTUS"
addExchangeCode "SNFE"
addExchangeCode "SOFFEX"
addExchangeCode "SWB"
addExchangeCode "SWX"

addExchangeCode "TPLUS1"
addExchangeCode "TPLUS2"
addExchangeCode "TRACKECN"
addExchangeCode "TRQXEN"
addExchangeCode "TSE"
addExchangeCode "TSE.JPN"

addExchangeCode "VALUE"
addExchangeCode "VENTURE"
addExchangeCode "VIRTX"
addExchangeCode "VSE"
addExchangeCode "VWAP"

ReDim Preserve mExchangeCodes(mMaxExchangeCodesIndex) As String
SortStrings mExchangeCodes

mExchangeCodesInitialised = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupCurrencyDescs()
Const ProcName As String = "setupCurrencyDescs"
On Error GoTo Err

Set mCurrencyDescsColl = CreateSortedDictionary(KeyTypeString)

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

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
