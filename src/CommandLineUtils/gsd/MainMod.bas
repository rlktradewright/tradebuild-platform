Attribute VB_Name = "MainMod"
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

Public Const ProjectName                            As String = "Gsd27"
Private Const ModuleName                            As String = "MainMod"

Private Const DefaultClientId                       As Long = 323115649

Private Const AsyncCommand                          As String = "ASYNC"
Private Const ExitCommand                           As String = "EXIT"
Private Const HelpCommand                           As String = "HELP"
Private Const Help1Command                          As String = "?"
Private Const InFileCommand                         As String = "INFILE"
Private Const OutpuPathCommand                      As String = "OUTPUTPATH"
Private Const ScanCommand                           As String = "SCAN"
Private Const SetParamCommand                       As String = "SETPARAM"
Private Const DateTimeFormatCommand                 As String = "DATETIMEFORMAT"

Private Const SwitchApiMessageLogging               As String = "APIMESSAGELOGGING"
Private Const SwitchCommandSeparator                As String = "SEP"
Private Const SwitchOutputPath                      As String = "outputpath"
Private Const SwitchTws                             As String = "TWS"


Private Const FilenameCharsPattern                  As String = "^[^/\*\?""<>|]*$"
Private Const OutputPathPattern                     As String = "(?:{(\$\w*)})"

Private Const NowVariable                           As String = "$NOW"
Private Const ParametersVariable                    As String = "$PARAMETERS"
Private Const ScancodeVariable                      As String = "$SCANCODE"
Private Const TimestampVariable                     As String = "$TIMESTAMP"
Private Const TodayVariable                         As String = "$TODAY"
Private Const NumberOfRowsVariable                  As String = "$NUMBEROFROWS"
Private Const InstrumentVariable                    As String = "$INSTRUMENT"
Private Const LocationCodeVariable                  As String = "$LOCATIONCODE"
Private Const AbovePriceVariable                    As String = "$ABOVEPRICE"
Private Const BelowPriceVariable                    As String = "$BELOWPRICE"
Private Const AboveVolumeVariable                   As String = "$ABOVEVOLUME"
Private Const AverageOptionVolumeAboveVariable      As String = "$AVERAGEOPTIONVOLUMEABOVE"
Private Const MarketCapAboveVariable                As String = "$MARKETCAPABOVE"
Private Const MarketCapBelowVariable                As String = "$MARKETCAPBELOW"
Private Const MoodyRatingAboveVariable              As String = "$MOODYRATINGABOVE"
Private Const MoodyRatingBelowVariable              As String = "$MOODYRATINGBELOW"
Private Const SpRatingAboveVariable                 As String = "$SPRATINGABOVE"
Private Const SpRatingBelowVariable                 As String = "$SPRATINGBELOW"
Private Const MaturityDateAboveVariable             As String = "$MATURITYDATEABOVE"
Private Const MaturityDateBelowVariable             As String = "$MATURITYDATEBELOW"
Private Const CouponRateAboveVariable               As String = "$COUPONRATEABOVE"
Private Const CouponRateBelowVariable               As String = "$COUPONRATEBELOW"
Private Const ExcludeConvertibleVariable            As String = "$EXCLUDECONVERTIBLE"
Private Const ScannerSettingPairsVariable           As String = "$SCANNERSETTINGPAIRS"
Private Const StockTypeFilterVariable               As String = "$STOCKTYPEFILTER"


Private Const ContractVariable                      As String = "$CONTRACT"
Private Const SymbolVariable                        As String = "$SYMBOL"
Private Const LocalSymbolVariable                   As String = "$LOCALSYMBOL"
Private Const SecTypeVariable                       As String = "$SECTYPE"
Private Const SecTypeAbbrvVariable                  As String = "$SECTYPEABBRV"
Private Const ExchangeVariable                      As String = "$EXCHANGE"
Private Const ExpiryVariable                        As String = "$EXPIRY"
Private Const CurrencyVariable                      As String = "$CURRENCY"
Private Const MultiplierVariable                    As String = "$MULTIPLIER"
Private Const StrikeVariable                        As String = "$STRIKE"
Private Const RightVariable                         As String = "$RIGHT"
Private Const RankVariable                          As String = "$RANK"
Private Const BenchmarkVariable                     As String = "$BENCHMARK"
Private Const DistanceVariable                      As String = "$DISTANCE"
Private Const LegsVariable                          As String = "$LEGS"
Private Const ProjectionVariable                    As String = "$PROJECTION"

Private Const DateFormatRawParameter                As String = "RAW"
Private Const DateFormatISOParameter                As String = "ISO"
Private Const DateFormatLocalParameter              As String = "LOCAL"

Private Const AppendOperator                        As String = ">>"
Private Const OverwriteOperator                     As String = ">"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mClp                                        As CommandLineParser
Private mCommandSeparator                           As String

Private mScanParams                                 As Parameters

Private mProcessors                                 As New EnumerableCollection

Private mTwsClient                                  As Client
Private mClientId                                   As Long
Private mProviderReady                              As Boolean

Private mAsync                                      As Boolean

Private mFilenameSubstitutionVariables()            As String
Private mMaxFilenameVariablesIndex                  As Long

Private mResultSubstitutionVariables()              As String
Private mMaxResultVariablesIndex                    As Long

Private mOutputPath                                 As String

Private mHistDataStore                              As HistoricalDataStore

Private mTimestampFormat                            As TimestampFormats
Private mTimestampDateOnlyFormat                    As TimestampFormats
Private mTimestampTimeOnlyFormat                    As TimestampFormats

Private mScanResultFormat                           As String

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

Public Property Get gFileSystemObject() As FileSystemObject
Static sFSO As FileSystemObject
If sFSO Is Nothing Then Set sFSO = New FileSystemObject
Set gFileSystemObject = sFSO
End Property

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("gsd", ProjectName)
Set gLogger = sLogger
End Property

Public Property Get gRegExp() As RegExp
Static sRegexp As RegExp
If sRegexp Is Nothing Then Set sRegexp = New RegExp
Set gRegExp = sRegexp
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateOutputStream( _
                ByVal pOutputPath As String, _
                ByVal pOutputFilename As String, _
                ByVal pProcessor As ScanProcessor, _
                ByVal pAppend As Boolean, _
                ByRef pMessage As String) As TextStream
Const ProcName As String = "gCreateOutputStream"
On Error GoTo Err

pOutputPath = gPerformFilenameVariableSubstitution(pOutputPath, pProcessor)
pOutputFilename = gPerformFilenameVariableSubstitution(pOutputFilename, pProcessor)
Dim lFilename As String
lFilename = gFileSystemObject.BuildPath(pOutputPath, pOutputFilename)
Set gCreateOutputStream = CreateWriteableTextFile( _
                                    lFilename, _
                                    pOverwrite:=(Not pAppend), _
                                    pUnicode:=True)

Exit Function

Err:
If Err.Number = 52 Then
    pMessage = "Invalid filename: " & lFilename
ElseIf Err.Number = ErrorCodes.ErrSecurityException Then
    pMessage = "Couldn't create file: " & lFilename & ": " & Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Public Function gGetContractName(ByVal pContractSpec As IContractSpecifier) As String
AssertArgument Not pContractSpec Is Nothing
With pContractSpec
    gGetContractName = .LocalSymbol & _
                        IIf(.Exchange <> "", "@" & .Exchange, "") & _
                        IIf(.CurrencyCode <> "", "(" & .CurrencyCode & ")", "")
End With
End Function

Public Sub gHandleFatalError(ev As ErrorEventData)
On Error Resume Next    ' ignore any further errors that might arise

gCon.WriteErrorString "Error "
gCon.WriteErrorString CStr(ev.ErrorCode)
gCon.WriteErrorString ": "
gCon.WriteErrorLine ev.ErrorMessage
gCon.WriteErrorLine "At:"
gCon.WriteErrorLine ev.ErrorSource
End Sub

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

Public Sub gNotifyAPIConnectionStateChange( _
                ByVal pState As ApiConnectionStates, _
                ByVal pMessage As String)
Const ProcName As String = "gNotifyAPIConnectionStateChange"
On Error GoTo Err

Select Case pState
Case ApiConnNotConnected
    gWriteLineToConsole "Not connected to TWS: " & pMessage
    mProviderReady = False
Case ApiConnConnecting
    gWriteLineToConsole "Connecting to TWS: " & pMessage
    mProviderReady = False
Case ApiConnConnected
    gWriteLineToConsole "Connected to TWS: " & pMessage
    mProviderReady = True
Case ApiConnFailed
    gWriteLineToConsole "Connection to TWS failed: " & pMessage
    mProviderReady = False
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gNotifyIBServerConnectionClosed()
gWriteLineToConsole "Connection from TWS to IB servers closed"
mProviderReady = False
End Sub

Public Sub gNotifyIBServerConnectionRecovered( _
                ByVal pDataLost As Boolean)
gWriteLineToConsole "Connection from TWS to IB servers recovered"
mProviderReady = True
End Sub

Public Sub gNotifyProcessorCompleted( _
                ByVal pProcessor As ScanProcessor, _
                ByVal pMessage As String)
Const ProcName As String = "gNotifyProcessorCompleted"
On Error GoTo Err

gWriteLineToConsole "Scan completed: " & _
                    pProcessor.ScanName & _
                    "; " & pMessage

If mProcessors.Contains(pProcessor) Then mProcessors.Remove pProcessor

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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

Public Sub gOutputScanItemToConsole()
Const ProcName As String = "gOutputScanItemToConsole"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gPerformFilenameVariableSubstitution( _
                ByVal pString As String, _
                ByVal pProcessor As ScanProcessor) As String
Const ProcName As String = "gPerformFilenameVariableSubstitution"
On Error GoTo Err

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = OutputPathPattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pString)

Dim s As String
Dim lCurrPosn As Long: lCurrPosn = 1

Dim lMatch As Match
For Each lMatch In lMatches
    s = s & Mid$(pString, lCurrPosn, lMatch.FirstIndex - lCurrPosn + 1)
    lCurrPosn = lMatch.FirstIndex + lMatch.Length + 1
    
    Dim r As String: r = ""
    
    Dim lVariable As String: lVariable = UCase$(lMatch.SubMatches(0))
    Select Case lVariable
    Case ScancodeVariable
        r = pProcessor.ScanName
    Case ParametersVariable
        Dim lParam As Parameter
        For Each lParam In pProcessor.ScanParameters
            If lParam.Value <> "" Then r = r & lParam.Name & "=" & lParam.Value & ";"
        Next
    Case TodayVariable
        r = FormatTimestamp(Now, mTimestampDateOnlyFormat)
    Case NowVariable
        r = FormatTimestamp(Now, mTimestampTimeOnlyFormat)
    Case TimestampVariable
        r = FormatTimestamp(Now, mTimestampFormat)
    Case NumberOfRowsVariable
        r = getScanParameterVariable(NumberOfRowsVariable, pProcessor)
    Case InstrumentVariable
        r = getScanParameterVariable(InstrumentVariable, pProcessor)
    Case LocationCodeVariable
        r = getScanParameterVariable(LocationCodeVariable, pProcessor)
    Case AbovePriceVariable
        r = getScanParameterVariable(AbovePriceVariable, pProcessor)
    Case BelowPriceVariable
        r = getScanParameterVariable(BelowPriceVariable, pProcessor)
    Case AboveVolumeVariable
        r = getScanParameterVariable(AboveVolumeVariable, pProcessor)
    Case AverageOptionVolumeAboveVariable
        r = getScanParameterVariable(AverageOptionVolumeAboveVariable, pProcessor)
    Case MarketCapAboveVariable
        r = getScanParameterVariable(MarketCapAboveVariable, pProcessor)
    Case MarketCapBelowVariable
        r = getScanParameterVariable(MarketCapBelowVariable, pProcessor)
    Case MoodyRatingAboveVariable
        r = getScanParameterVariable(MoodyRatingAboveVariable, pProcessor)
    Case MoodyRatingBelowVariable
        r = getScanParameterVariable(MoodyRatingBelowVariable, pProcessor)
    Case SpRatingAboveVariable
        r = getScanParameterVariable(SpRatingAboveVariable, pProcessor)
    Case SpRatingBelowVariable
        r = getScanParameterVariable(SpRatingBelowVariable, pProcessor)
    Case MaturityDateAboveVariable
        r = getScanParameterVariable(MaturityDateAboveVariable, pProcessor)
    Case MaturityDateBelowVariable
        r = getScanParameterVariable(MaturityDateBelowVariable, pProcessor)
    Case CouponRateAboveVariable
        r = getScanParameterVariable(CouponRateAboveVariable, pProcessor)
    Case CouponRateBelowVariable
        r = getScanParameterVariable(CouponRateBelowVariable, pProcessor)
    Case ExcludeConvertibleVariable
        r = getScanParameterVariable(ExcludeConvertibleVariable, pProcessor)
    Case ScannerSettingPairsVariable
        r = getScanParameterVariable(ScannerSettingPairsVariable, pProcessor)
    Case StockTypeFilterVariable
        r = getScanParameterVariable(StockTypeFilterVariable, pProcessor)
    Case Default
        Assert False, "Unexpected substitution variable: " & lVariable
    End Select
    s = s & escapeNonFilenameChars(r)
Next

gPerformFilenameVariableSubstitution = s & Right$(pString, Len(pString) - lCurrPosn + 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPerformResultVariableSubstitution( _
                ByVal pString As String, _
                ByVal pScanResult As IScanResult) As String
Const ProcName As String = "gPerformResultVariableSubstitution"
On Error GoTo Err

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = OutputPathPattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pString)

Dim lContractSpec As IContractSpecifier
Set lContractSpec = pScanResult.Contract.Specifier

Dim s As String
Dim lCurrPosn As Long: lCurrPosn = 1

Dim lMatch As Match
For Each lMatch In lMatches
    s = s & Mid$(pString, lCurrPosn, lMatch.FirstIndex - lCurrPosn + 1)
    lCurrPosn = lMatch.FirstIndex + lMatch.Length + 1
    
    Dim r As String
    
    Dim lVariable As String: lVariable = UCase$(lMatch.SubMatches(0))
    Select Case lVariable
    Case ContractVariable
        r = gGetContractName(lContractSpec)
    Case SymbolVariable
        r = lContractSpec.Symbol
    Case LocalSymbolVariable
        r = lContractSpec.LocalSymbol
    Case SecTypeVariable
        r = SecTypeToString(lContractSpec.SecType)
    Case SecTypeAbbrvVariable
        r = SecTypeToShortString(lContractSpec.SecType)
    Case ExchangeVariable
        r = lContractSpec.Exchange
    Case ExpiryVariable
        r = lContractSpec.Expiry
    Case CurrencyVariable
        r = lContractSpec.CurrencyCode
    Case MultiplierVariable
        r = lContractSpec.Multiplier
    Case StrikeVariable
        r = lContractSpec.Strike
    Case RightVariable
        r = OptionRightToString(lContractSpec.Right)
    Case RankVariable
        r = pScanResult.Rank
    Case BenchmarkVariable
        r = pScanResult.Attributes.GetParameterValue("Benchmark", "")
    Case DistanceVariable
        r = pScanResult.Attributes.GetParameterValue("Distance", "")
    Case LegsVariable
        r = pScanResult.Attributes.GetParameterValue("Legs", "")
    Case ProjectionVariable
        r = pScanResult.Attributes.GetParameterValue("Projection", "")
    Case Default
        Assert False, "Unexpected substitution variable: " & lVariable
    End Select
    s = s & escapeNonFilenameChars(r)
Next

gPerformResultVariableSubstitution = s & Right$(pString, Len(pString) - lCurrPosn + 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gWriteErrorLine( _
                ByVal pMessage As String)
Const ProcName As String = "gWriteErrorLine"

Dim s As String
s = "Error: " & pMessage
gCon.WriteErrorLine s
LogMessage "StdErr: " & s
End Sub

Public Sub gWriteLineToConsole(ByVal pMessage As String, Optional ByVal pLogit As Boolean)
Const ProcName As String = "gWriteLineToConsole"

If pLogit Then LogMessage "Con: " & pMessage
gCon.WriteLineToConsole pMessage
End Sub

Public Sub gWriteLineToStdOut(ByVal pMessage As String)
Const ProcName As String = "gWriteLineToStdOut"

LogMessage "StdOut: " & pMessage
gCon.WriteLine pMessage
End Sub

Public Sub Main()
Const ProcName As String = "Main"
On Error GoTo Err

Set gCon = GetConsole

If Trim$(Command) = "/?" Or Trim$(Command) = "-?" Then
    showUsage
    Exit Sub
End If

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "gsd"
SetupDefaultLogging Command, True, True

logProgramId

setupSubstitutionVariables

mScanResultFormat = "{" & SecTypeAbbrvVariable & "}:{" & ContractVariable & "}"

mTimestampFormat = TimestampDateAndTimeISO8601
mTimestampDateOnlyFormat = TimestampDateOnlyISO8601
mTimestampTimeOnlyFormat = TimestampTimeOnlyISO8601

mCommandSeparator = ";"

Set mClp = CreateCommandLineParser(Command)

Dim lLogApiMessages As ApiMessageLoggingOptions
Dim lLogRawApiMessages As ApiMessageLoggingOptions
Dim lLogApiMessageStats As Boolean
If Not validateApiMessageLogging( _
                mClp.SwitchValue(SwitchApiMessageLogging), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats) Then
    gWriteLineToConsole "API message logging setting is invalid", True
    Exit Sub
End If

If mClp.Switch(SwitchCommandSeparator) Then mCommandSeparator = mClp.SwitchValue(SwitchCommandSeparator)
Assert Len(mCommandSeparator) = 1, "The command separator must be a single character"

If mClp.Switch(SwitchOutputPath) Then processOutputPathCommand mClp.SwitchValue(SwitchOutputPath)

If Not setupTwsApi(mClp.SwitchValue(SwitchTws), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats) Then
    showUsage
    Exit Sub
End If

process

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addFilenameSubstitutionVariable(ByVal pVariable As String)
Const ProcName As String = "addFilenameSubstitutionVariable"
On Error GoTo Err

mMaxFilenameVariablesIndex = mMaxFilenameVariablesIndex + 1
If mMaxFilenameVariablesIndex > UBound(mFilenameSubstitutionVariables) Then
    ReDim Preserve mFilenameSubstitutionVariables(2 * (UBound(mFilenameSubstitutionVariables) + 1) - 1) As String
End If
mFilenameSubstitutionVariables(mMaxFilenameVariablesIndex) = UCase$(pVariable)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addResultSubstitutionVariable(ByVal pVariable As String)
Const ProcName As String = "addResultSubstitutionVariable"
On Error GoTo Err

mMaxResultVariablesIndex = mMaxResultVariablesIndex + 1
If mMaxResultVariablesIndex > UBound(mResultSubstitutionVariables) Then
    ReDim Preserve mResultSubstitutionVariables(2 * (UBound(mResultSubstitutionVariables) + 1) - 1) As String
End If
mResultSubstitutionVariables(mMaxResultVariablesIndex) = UCase$(pVariable)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function escapeNonFilenameChars(ByVal pFilename As String) As String
Dim ar() As Byte
ar = pFilename
Dim i As Long
For i = 0 To UBound(ar) - 1
    Dim c As String: c = Chr$(ar(i))
    If c = "/" Then
        c = "-"
    ElseIf c = ":" Then
        c = "."
    ElseIf c = "*" Then
        c = "'"
    ElseIf c = "?" Then
        c = "~"
    ElseIf c = """" Then
        c = "'"
    ElseIf c = "<" Then
        c = "_"
    ElseIf c = ">" Then
        c = "_"
    ElseIf c = "|" Then
        c = "^"
    End If
    ar(i) = Asc(c)
Next
escapeNonFilenameChars = ar
End Function

Private Function getInputLine() As String
Const ProcName As String = "getInputLine"
On Error GoTo Err

getInputLine = Trim$(gCon.ReadLine(":"))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getScanParameterVariable( _
                ByVal pVariableName As String, _
                ByVal pProcessor As ScanProcessor) As String
getScanParameterVariable = pProcessor.ScanParameters.GetParameterValue(Mid$(pVariableName, 2), "")
End Function

Private Function isValidPath(ByVal pPath As String) As Boolean
Const ProcName As String = "isValidPath"
On Error GoTo Err

isValidPath = True

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = FilenameCharsPattern
If Not lRegExp.Test(pPath) Then
    gWriteErrorLine "Invalid characters: path cannot contain  / * ? "" < > | "
    isValidPath = False
    Exit Function
End If

lRegExp.Pattern = OutputPathPattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pPath)

Dim lMatch As Match
For Each lMatch In lMatches
    Dim lVariable As String: lVariable = lMatch.SubMatches(0)
    If Not isValidFilenameSubstitutionVariable(lVariable) Then
        gWriteErrorLine lVariable & " is not a valid substitution variable"
        isValidPath = False
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidFilenameSubstitutionVariable(ByVal pString As String) As Boolean
isValidFilenameSubstitutionVariable = BinarySearchStrings( _
                                UCase$(pString), _
                                mFilenameSubstitutionVariables, _
                                0, _
                                mMaxFilenameVariablesIndex + 1) >= 0
End Function

Private Function isValidResultSubstitutionVariable(ByVal pString As String) As Boolean
isValidResultSubstitutionVariable = BinarySearchStrings( _
                                UCase$(pString), _
                                mResultSubstitutionVariables, _
                                0, _
                                mMaxResultVariablesIndex + 1) >= 0
End Function

Private Sub logProgramId()
Const ProcName As String = "logProgramId"
On Error GoTo Err

Dim s As String
s = App.ProductName & _
    " V" & _
    App.Major & _
    "." & App.Minor & _
    "." & App.Revision & _
    IIf(App.FileDescription <> "", "." & App.FileDescription, "") & _
    vbCrLf & _
    App.LegalCopyright
gWriteLineToConsole s, False
s = s & vbCrLf & "Arguments: " & Command
LogMessage s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

Set mScanParams = New Parameters

Dim lContinue As Boolean
processCommandLineCommands lContinue

If lContinue Then
    processStdInComands
End If

Do While mProcessors.Count <> 0
    Wait 50
Loop

If Not mTwsClient Is Nothing Then
    gWriteLineToConsole "Releasing API connection", True
    mTwsClient.Finish
    ' allow time for the socket connection to be nicely released
    Wait 10
End If

gWriteLineToConsole "Exiting", True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processAsyncCommand(ByVal pParams As String)
pParams = UCase$(pParams)
If pParams = "" Or pParams = "YES" Or pParams = "TRUE" Or pParams = "ON" Then
    mAsync = True
ElseIf pParams = "NO" Or pParams = "FALSE" Or pParams = "OFF" Then
    mAsync = False
Else
    gWriteErrorLine "parameter must be YES, NO, ON, OFF, TRUE or FALSE"
End If
End Sub

Private Sub processCommand(ByVal pCommandString As String)
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim lCommand As String
lCommand = UCase$(Split(pCommandString, " ")(0))

Dim params As String
params = Trim$(Right$(pCommandString, Len(pCommandString) - Len(lCommand)))

Select Case lCommand
Case ScanCommand
    processScanCommand params
Case SetParamCommand
    processSetParamCommand params
Case HelpCommand, Help1Command
    showStdInHelp
Case OutpuPathCommand
    processOutputPathCommand params
Case AsyncCommand
    processAsyncCommand params
Case DateTimeFormatCommand
    processDateTimeFormatCommand params
Case InFileCommand
    processInfileCommand params
Case Else
    gCon.WriteErrorLine "Invalid command '" & lCommand & "'"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCommandLineCommands( _
                ByRef pContinue As Boolean)
Const ProcName As String = "processCommandLineCommands"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(mClp.Arg(0), mCommandSeparator)

If lClp.NumberOfArgs = 0 Then
    pContinue = True
    Exit Sub
End If

Dim lInputString As String

Dim i As Long
Do
    If mProviderReady Then
        i = i + 1
        
        If i > lClp.NumberOfArgs Then Exit Do
        
        lInputString = lClp.Arg(i - 1)
        If UCase$(lInputString) = ExitCommand Then
            pContinue = False
            Exit Sub
        End If
        
        If lInputString = "" Then
            ' ignore blank lines
        ElseIf Left$(lInputString, 1) = "#" Then
            LogMessage "cmd: " & lInputString
            ' ignore comments
        Else
            LogMessage "cmd: " & lInputString
            processCommand lInputString
        End If
    Else
        Wait 10
    End If
Loop

pContinue = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processDateTimeFormatCommand( _
                ByVal params As String)
Const ProcName As String = "processDateTimeFormatCommand"
On Error GoTo Err

params = UCase$(params)

If params = DateFormatISOParameter Then
    mTimestampFormat = TimestampDateAndTimeISO8601
    mTimestampDateOnlyFormat = TimestampDateOnlyISO8601
    mTimestampTimeOnlyFormat = TimestampTimeOnlyISO8601
ElseIf params = DateFormatLocalParameter Then
    mTimestampFormat = TimestampDateAndTimeLocal
    mTimestampDateOnlyFormat = TimestampDateOnlyLocal
    mTimestampTimeOnlyFormat = TimestampTimeOnlyLocal
ElseIf params = DateFormatRawParameter Then
    mTimestampFormat = TimestampDateAndTime
    mTimestampDateOnlyFormat = TimestampDateOnly
    mTimestampTimeOnlyFormat = TimestampTimeOnly
Else
    gWriteErrorLine "Invalid date/time format '" & params & "'"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processInfileCommand( _
                ByVal params As String)
Const ProcName As String = "processInfileCommand"
On Error GoTo Err

Dim lTs As TextStream
Set lTs = gFileSystemObject.OpenTextFile(params, ForReading)

Do Until lTs.AtEndOfStream
    If mProviderReady Then
        Dim lInputString As String
        lInputString = lTs.ReadLine
        
        If UCase$(lInputString) = ExitCommand Then Exit Sub
        
        If lInputString = "" Then
            ' ignore blank lines
        ElseIf Left$(lInputString, 1) = "#" Then
            LogMessage "file: " & lInputString
            ' ignore comments
        Else
            LogMessage "file: " & lInputString
            processCommand lInputString
        End If
    End If
Loop

Exit Sub

Err:
If Err.Number = 52 Then
    gWriteErrorLine "Invalid filename: " & params
ElseIf Err.Number = VBErrorCodes.VbErrFileNotFound Then
    gWriteErrorLine "File does not exist: " & params
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processOutputPathCommand( _
                ByVal params As String)
Const ProcName As String = "processOutputPathCommand"
On Error GoTo Err

If isValidPath(params) Then mOutputPath = params

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processScanCommand(ByVal pParams As String)
Const ProcName As String = "processScanCommand"
On Error GoTo Err

If Not mAsync And mProcessors.Count <> 0 Then
    gWriteErrorLine "Cannot start - already running"
    Exit Sub
End If

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pParams)

Dim lScanCode As String: lScanCode = lClp.Arg(0)
If lScanCode = "" Then
    gWriteErrorLine "Scan code must be the first parameter"
    Exit Sub
End If

Dim lPathAndFilename As String

Dim lAppend As Boolean
If lClp.NumberOfArgs = 1 Then
    lAppend = False
ElseIf lClp.NumberOfArgs = 2 Then
    If lClp.Arg(1) = AppendOperator Then
        lAppend = True
    ElseIf lClp.Arg(1) = OverwriteOperator Then
        lAppend = False
    Else
        mScanParams.SetParameterValues lClp.Arg(1)
    End If
ElseIf lClp.NumberOfArgs = 3 Then
    If lClp.Arg(1) = AppendOperator Then
        lAppend = True
        lPathAndFilename = lClp.Arg(2)
    ElseIf lClp.Arg(1) = OverwriteOperator Then
        lAppend = False
        lPathAndFilename = lClp.Arg(2)
    Else
        mScanParams.SetParameterValues lClp.Arg(1)
        If lClp.Arg(2) = AppendOperator Then
            lAppend = True
        ElseIf lClp.Arg(2) = OverwriteOperator Then
            lAppend = False
        Else
            lAppend = False
            lPathAndFilename = lClp.Arg(2)
        End If
    End If
ElseIf lClp.NumberOfArgs = 4 Then
    mScanParams.SetParameterValues lClp.Arg(1)
    lPathAndFilename = lClp.Arg(3)
    If lClp.Arg(2) = AppendOperator Then
        lAppend = True
    ElseIf lClp.Arg(2) = OverwriteOperator Then
        lAppend = False
    Else
        gWriteErrorLine "Third argument must be '>' or '>>'"
        Exit Sub
    End If
ElseIf lClp.Arg(0) = AppendOperator Then
    lAppend = True
ElseIf lClp.Arg(0) = OverwriteOperator Then
    lAppend = False
Else
    gWriteErrorLine "Too many arguments"
    Exit Sub
End If

If Not isValidPath(lPathAndFilename) Then Exit Sub

Dim lScanProcessor As ScanProcessor: Set lScanProcessor = New ScanProcessor
mProcessors.Add lScanProcessor



'mScanParams.SetParameterValue "NumberOfRows", "10"
'mScanParams.SetParameterValue "Instrument", "STK"
'mScanParams.SetParameterValue "AbovePrice", "3.0"
'mScanParams.SetParameterValue "MarketCapAbove", "100000000"
'mScanParams.SetParameterValue "ScannerSettingPairs", "Annual,true"
'mScanParams.SetParameterValue "LocationCode", "STK.ARCA"

lScanProcessor.Scan mHistDataStore, _
                    lScanCode, _
                    mScanParams, _
                    Nothing, _
                    Nothing, _
                    False, _
                    mScanResultFormat, _
                    mOutputPath, _
                    lPathAndFilename, _
                    lAppend

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSetParamCommand( _
                ByVal pParams As String)
Const ProcName As String = "processSetParamCommand"
On Error GoTo Err

mScanParams.SetParameterValues pParams

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStdInComands()
Const ProcName As String = "processStdInComands"
On Error GoTo Err

Do
    Dim lInputString As String
    lInputString = Trim$(gCon.ReadLine(":"))
    If lInputString = gCon.EofString Or UCase$(lInputString) = ExitCommand Then Exit Do
        
    If lInputString = "" Then
        ' ignore blank lines
    ElseIf Left$(lInputString, 1) = "#" Then
        LogMessage "con: " & lInputString
        ' ignore comments
    ElseIf mProviderReady Then
        LogMessage "con: " & lInputString
        processCommand lInputString
    Else
        gWriteErrorLine "Not ready"
        Wait 10
    End If
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setupSubstitutionVariables()
ReDim mFilenameSubstitutionVariables(15) As String
mMaxFilenameVariablesIndex = -1

addFilenameSubstitutionVariable ScancodeVariable
addFilenameSubstitutionVariable NowVariable
addFilenameSubstitutionVariable ParametersVariable
addFilenameSubstitutionVariable TimestampVariable
addFilenameSubstitutionVariable TodayVariable
addFilenameSubstitutionVariable NumberOfRowsVariable
addFilenameSubstitutionVariable InstrumentVariable
addFilenameSubstitutionVariable LocationCodeVariable
addFilenameSubstitutionVariable AbovePriceVariable
addFilenameSubstitutionVariable BelowPriceVariable
addFilenameSubstitutionVariable AboveVolumeVariable
addFilenameSubstitutionVariable AverageOptionVolumeAboveVariable
addFilenameSubstitutionVariable MarketCapAboveVariable
addFilenameSubstitutionVariable MarketCapBelowVariable
addFilenameSubstitutionVariable MoodyRatingAboveVariable
addFilenameSubstitutionVariable MoodyRatingBelowVariable
addFilenameSubstitutionVariable SpRatingAboveVariable
addFilenameSubstitutionVariable SpRatingBelowVariable
addFilenameSubstitutionVariable MaturityDateAboveVariable
addFilenameSubstitutionVariable MaturityDateBelowVariable
addFilenameSubstitutionVariable CouponRateAboveVariable
addFilenameSubstitutionVariable CouponRateBelowVariable
addFilenameSubstitutionVariable ExcludeConvertibleVariable
addFilenameSubstitutionVariable ScannerSettingPairsVariable
addFilenameSubstitutionVariable StockTypeFilterVariable

SortStrings mFilenameSubstitutionVariables, EndIndex:=mMaxFilenameVariablesIndex


ReDim mResultSubstitutionVariables(15) As String
mMaxResultVariablesIndex = -1

addResultSubstitutionVariable ContractVariable
addResultSubstitutionVariable SymbolVariable
addResultSubstitutionVariable LocalSymbolVariable
addResultSubstitutionVariable SecTypeVariable
addResultSubstitutionVariable SecTypeAbbrvVariable
addResultSubstitutionVariable ExchangeVariable
addResultSubstitutionVariable ExpiryVariable
addResultSubstitutionVariable CurrencyVariable
addResultSubstitutionVariable MultiplierVariable
addResultSubstitutionVariable StrikeVariable
addResultSubstitutionVariable RightVariable
addResultSubstitutionVariable RankVariable
addResultSubstitutionVariable BenchmarkVariable
addResultSubstitutionVariable DistanceVariable
addResultSubstitutionVariable LegsVariable
addResultSubstitutionVariable ProjectionVariable

SortStrings mResultSubstitutionVariables, EndIndex:=mMaxResultVariablesIndex
End Sub

Private Function setupTwsApi( _
                ByVal SwitchValue As String, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean) As Boolean
Const ProcName As String = "setupTwsApi"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(SwitchValue, ",")

Dim server As String
server = lClp.Arg(0)

Dim port As String
port = lClp.Arg(1)
If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gWriteErrorLine "port must be an integer > 0"
    setupTwsApi = False
End If
    
Dim clientId As String
clientId = lClp.Arg(2)
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "clientId must be an integer >= 0 and <= 999999999"
    setupTwsApi = False
Else
    mClientId = CLng(clientId)
End If

Dim connectionRetryInterval As String
connectionRetryInterval = lClp.Arg(3)
If connectionRetryInterval = "" Then
ElseIf Not IsInteger(connectionRetryInterval, 0, 3600) Then
    gWriteErrorLine "Error: connection retry interval must be an integer >= 0 and <= 3600"
    setupTwsApi = False
End If

Dim lListener As New TWSConnectionMonitor

If connectionRetryInterval = "" Then
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            mClientId, _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats, _
                            pConnectionStateListener:=lListener)
Else
    Set mTwsClient = GetClient(server, _
                            CLng(port), _
                            mClientId, _
                            pConnectionRetryIntervalSecs:=CLng(connectionRetryInterval), _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats, _
                            pConnectionStateListener:=lListener)
End If

Set mHistDataStore = mTwsClient.GetHistoricalDataStore

setupTwsApi = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showStdInHelp()
gWriteLineToConsole "StdIn Format:"
gWriteLineToConsole ""
gWriteLineToConsole "#comment"
gWriteLineToConsole ""
gWriteLineToConsole "???????????? To be supplied"
End Sub

Private Sub showUsage()
gWriteLineToConsole "Usage:"
gWriteLineToConsole "gsd27 -tws[:[<twsserver>][,[<port>][,[<clientid>]]]]"
gWriteLineToConsole "       [-scan:scanname]"
gWriteLineToConsole "       [-resultsdir:<resultspath>] [-log:<logfilepath>]"
gWriteLineToConsole "       [-loglevel:[ I | N | D | M | H }]"
gWriteLineToConsole "       [-apimessagelogging:[D|A|N][D|A|N][Y|N]]"
gWriteLineToConsole "       [-scopename:<scope>] [-recoveryfiledir:<recoverypath>]"
gWriteLineToConsole ""
gWriteLineToConsole "  where"
gWriteLineToConsole ""
gWriteLineToConsole "    twsserver  ::= STRING name or IP address of computer where TWS/Gateway"
gWriteLineToConsole "                          is running"
gWriteLineToConsole "    port       ::= INTEGER port to be used for API connection"
gWriteLineToConsole "    clientid   ::= INTEGER client id >=0 to be used for API connection (default"
gWriteLineToConsole "                           value is " & DefaultClientId & ")"
gWriteLineToConsole "    resultspath ::= path to the folder in which results files are to be created"
gWriteLineToConsole "                    (defaults to the logfile path)"
gWriteLineToConsole "    logfilepath ::= path to the folder where the program logfile is to be"
gWriteLineToConsole "                    created"
gWriteLineToConsole "    scanname       ::= name of the market scan to run"
gWriteLineToConsole "    recoverypath ::= path to the folder in which order recovery files are to"
gWriteLineToConsole "                     be created(defaults to the logfile path)"
gWriteLineToConsole ""
End Sub

Private Function validateApiMessageLogging( _
                ByVal pApiMessageLogging As String, _
                ByRef pLogApiMessages As ApiMessageLoggingOptions, _
                ByRef pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByRef pLogApiMessageStats As Boolean) As Boolean
Const Always As String = "A"
Const Default As String = "D"
Const No As String = "N"
Const None As String = "N"
Const Yes As String = "Y"

pApiMessageLogging = UCase$(pApiMessageLogging)

validateApiMessageLogging = False
If Len(pApiMessageLogging) = 0 Then pApiMessageLogging = Default & Default & No
If Len(pApiMessageLogging) <> 3 Then Exit Function

Dim s As String
s = Left$(pApiMessageLogging, 1)
If s = None Then
    pLogApiMessages = ApiMessageLoggingOptionNone
ElseIf s = Default Then
    pLogApiMessages = ApiMessageLoggingOptionDefault
ElseIf s = Always Then
    pLogApiMessages = ApiMessageLoggingOptionAlways
Else
    Exit Function
End If

s = Mid(pApiMessageLogging, 2, 1)
If s = None Then
    pLogRawApiMessages = ApiMessageLoggingOptionNone
ElseIf s = Default Then
    pLogRawApiMessages = ApiMessageLoggingOptionDefault
ElseIf s = Always Then
    pLogRawApiMessages = ApiMessageLoggingOptionAlways
Else
    Exit Function
End If

s = Mid(pApiMessageLogging, 3, 1)
If s = No Then
    pLogApiMessageStats = False
ElseIf s = Yes Then
    pLogApiMessageStats = True
Else
    Exit Function
End If

validateApiMessageLogging = True
End Function




