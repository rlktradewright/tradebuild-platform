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

Public Enum DataSources
    FromDb
    FromFile
    FromTws
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "gbd"
Private Const ModuleName                            As String = "MainMod"

Private Const ContractCommand                       As String = "CONTRACT"
Private Const FromCommand                           As String = "FROM"
Private Const ToCommand                             As String = "TO"
Private Const StartCommand                          As String = "START"
Private Const StopCommand                           As String = "STOP"
Private Const NumberCommand                         As String = "NUMBER"
Private Const TimeframeCommand                      As String = "TIMEFRAME"
Private Const SessCommand                           As String = "SESS"
Private Const NonSessCommand                        As String = "NONSESS"
Private Const SessionOnlyCommand                    As String = "SESSIONONLY"
Private Const MillisecsCommand                      As String = "MILLISECS"
Private Const NoMillisecsCommand                    As String = "NOMILLISECS"
Private Const HelpCommand                           As String = "HELP"
Private Const Help1Command                          As String = "?"
Private Const SessionEndTimeCommand                 As String = "SESSIONENDTIME"
Private Const SessionStartTimeCommand               As String = "SESSIONSTARTTIME"
Private Const DateOnlyCommmand                      As String = "DATEONLY"
Private Const OutpuPathCommand                      As String = "OUTPUTPATH"
Private Const AsyncCommand                          As String = "ASYNC"
Private Const EntireSessionCommand                  As String = "ENTIRESESSION"
Private Const ExitCommand                           As String = "EXIT"
Private Const DateTimeFormatCommand                 As String = "DATETIMEFORMAT"
Private Const InFileCommand                         As String = "INFILE"
Private Const EchoCommand                           As String = "ECHO"
Private Const StartEchoCommand                      As String = "STARTECHO"
Private Const ResultFormatCommand                   As String = "RESULTFORMAT"
Private Const EchoResultFormatCommand               As String = "ECHORESULTFORMAT"

Private Const DateFormatRawParameter                As String = "RAW"
Private Const DateFormatISOParameter                As String = "ISO"
Private Const DateFormatLocalParameter              As String = "LOCAL"
Private Const OutputFormatTimestampParameter        As String = "T"
Private Const OutputFormatOpenParameter             As String = "O"
Private Const OutputFormatHighParameter             As String = "H"
Private Const OutputFormatLowParameter              As String = "L"
Private Const OutputFormatCloseParameter            As String = "C"
Private Const OutputFormatVolumeParameter           As String = "V"
Private Const OutputFormatTickVolumeParameter       As String = "TV"
Private Const OutputFormatOpenInterestParameter     As String = "OI"

Private Const AppendOperator                        As String = ">>"
Private Const OverwriteOperator                     As String = ">"

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
Private Const TodayVariable                         As String = "$TODAY"
Private Const YesterdayVariable                     As String = "$YESTERDAY"
Private Const FromDateVariable                      As String = "$FROMDATE"
Private Const FromDateTimeVariable                  As String = "$FROMDATETIME"
Private Const FromTimeVariable                      As String = "$FROMTIME"
Private Const ToDateVariable                        As String = "$TODATE"
Private Const ToDateTimeVariable                    As String = "$TODATETIME"
Private Const ToTimeVariable                        As String = "$TOTIME"
Private Const TimeframeVariable                     As String = "$TIMEFRAME"
Private Const NumberRequestedVariable               As String = "$NUMBERREQUESTED"
Private Const NumberReturnedVariable                As String = "$NUMBERRETURNED"
Private Const NewlineVariable                       As String = "$NEWLINE"

Private Const BarNumberVariable                     As String = "$BARNUMBER"
Private Const TimestampVariable                     As String = "$TIMESTAMP"
Private Const OpenVariable                          As String = "$OPEN"
Private Const HighVariable                          As String = "$HIGH"
Private Const LowVariable                           As String = "$LOW"
Private Const CloseVariable                         As String = "$CLOSE"
Private Const VolumeVariable                        As String = "$VOLUME"
Private Const TickVolumeVariable                    As String = "$TICKVOLUME"


Private Const SwitchCommandSeparator                As String = "SEP"
Private Const SwitchFromDb                          As String = "fromdb"
Private Const SwitchFromFile                        As String = "fromfile"
Private Const SwitchFromTws                         As String = "fromtws"
Private Const SwitchLogToConsole                    As String = "logtoconsole"
Private Const SwitchOutputPath                      As String = "outputpath"
Private Const SwitchApiMessageLogging               As String = "APIMESSAGELOGGING"

Private Const DefaultClientId                       As Long = 205644991

Private Const Time235900                            As Double = 0.99930556712963

Private Const FilenameCharsPattern                  As String = "^[^/\*\?""<>|]*$"
Private Const SubstitutionVariablePattern           As String = "(?:{(\$\w*)})"

Private Const DefaultResultFormat                   As String = "{" & TimestampVariable & "}," & _
                                                                "{" & OpenVariable & "}," & _
                                                                "{" & HighVariable & "}," & _
                                                                "{" & LowVariable & "}," & _
                                                                "{" & CloseVariable & "}," & _
                                                                "{" & VolumeVariable & "}," & _
                                                                "{" & TickVolumeVariable & "}"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console

Public gLogToConsole                                As Boolean

Private mClp                                        As CommandLineParser

Private mCommandSeparator                           As String

Private mAsync                                      As Boolean
Private mEntireSession                              As Boolean

Private mProcessors                                 As New EnumerableCollection
Private mCurrentProcessor                           As IProcessor

Private mDataSource                                 As DataSources

Private mTickfileName                               As String

Private mLineNumber                                 As Long
Private mContractSpec                               As IContractSpecifier
Private mFrom                                       As Date
Private mSpecialFrom                                As String
Private mTo                                         As Date
Private mSpecialTo                                  As String
Private mNumber                                     As Long
Private mTimePeriod                                 As TimePeriod
Private mSessionOnly                                As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mIncludeMillisecs                           As Boolean

Private mHistDataStore                              As IHistoricalDataStore
Private mContractStore                              As IContractStore

Private mSessionEndTime                             As Date
Private mSessionStartTime                           As Date

Private mNormaliseDailyBarTimestamps                As Boolean

Private mTWSConnectionMonitor                       As TWSConnectionMonitor

Private mProviderReady                              As Boolean

Private mOutputPath                                 As String

Private mFilenameSubstitutionVariables()            As String
Private mMaxFilenameVariablesIndex                  As Long

Private mResultSubstitutionVariables()              As String
Private mMaxResultVariablesIndex                    As Long

Private mTimestampFormat                            As TimestampFormats
Private mTimestampDateOnlyFormat                    As TimestampFormats
Private mTimestampTimeOnlyFormat                    As TimestampFormats

Private mErrorStringBuilder                          As StringBuilder

Private mStartResultFormat                           As String
Private mStartEchoResultFormat                       As String

Private mEchoToStdOut                               As Boolean

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
If sLogger Is Nothing Then Set sLogger = CreateFormattingLogger("gbd", ProjectName)
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
                ByVal pProcessor As IProcessor, _
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
gGetContractName = pContractSpec.LocalSymbol & _
                    IIf(pContractSpec.Exchange <> "", "@" & pContractSpec.Exchange, "")
End Function

Public Sub gHandleFatalError(ev As ErrorEventData)
On Error Resume Next    ' ignore any further errors that might arise

gWriteErrorString "Error "
gWriteErrorString CStr(ev.ErrorCode)
gWriteErrorString ": "
gWriteErrorString ev.ErrorMessage
gWriteErrorString vbCrLf
gWriteErrorString "At:"
gWriteErrorString vbCrLf
gWriteErrorString ev.ErrorSource
gWriteErrorLine "Program exiting"
Set mCurrentProcessor = Nothing
EndProcess 1
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

Public Sub gNotifyDataRetrieved(ByVal pProcessor As IProcessor)
Const ProcName As String = "gNotifyDataRetrieved"
On Error GoTo Err

LogMessage "Data retrieved from source for contract: " & gGetContractName(pProcessor.ContractSpec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gNotifyFetchCancelled( _
                ByVal pProcessor As IProcessor)
Const ProcName As String = "gNotifyFetchCancelled"
On Error GoTo Err

If pProcessor.ContractSpec Is Nothing Then
    gWriteLineToConsole "Fetch cancelled for contract " & pProcessor.DataSourceName
Else
    gWriteLineToConsole "Fetch cancelled for contract " & gGetContractName(pProcessor.ContractSpec)
End If
If mProcessors.Contains(pProcessor) Then mProcessors.Remove pProcessor
Set mCurrentProcessor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gNotifyFetchCompleted( _
                ByVal pProcessor As IProcessor, _
                ByVal pDurationSecs As Double)
Const ProcName As String = "gNotifyFetchCompleted"
On Error GoTo Err

gWriteLineToConsole "Fetch completed in " & _
            Format(pDurationSecs, "0.0") & " secs: " & _
            pProcessor.NumberOfBarsOutput & _
            " bars for contract: " & _
            gGetContractName(pProcessor.ContractSpec)

If mProcessors.Contains(pProcessor) Then mProcessors.Remove pProcessor
Set mCurrentProcessor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gNotifyFetchFailed( _
                ByVal pProcessor As IProcessor, _
                ByVal pErrorMessage As String)
Const ProcName As String = "gNotifyFetchFailed"
On Error GoTo Err

gWriteLineToConsole "Fetch failed for " & pProcessor.DataSourceName & ": " & pErrorMessage
If mProcessors.Contains(pProcessor) Then mProcessors.Remove pProcessor
Set mCurrentProcessor = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gNotifyFetchStarted( _
                ByVal pProcessor As IProcessor)
Const ProcName As String = "gNotifyFetchStarted"
On Error GoTo Err

LogMessage "Fetch started for contract " & gGetContractName(pProcessor.ContractSpec)

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

Public Sub gOutputBarToStdOut( _
                ByVal pUseEchoResultFormat As Boolean, _
                ByVal pProcessor As IProcessor, _
                ByVal pBar As Bar, _
                ByVal pTimeframe As TimePeriod, _
                ByVal pBarnumber As Long, _
                ByVal pNormaliseDailyBarTimestamps As Boolean, _
                ByVal pIncludeMillisecs As Boolean, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double)
Const ProcName As String = "gOutputBarToStdOut"
On Error GoTo Err

If pBar Is Nothing Then Exit Sub

gWriteLineToStdOut formatBar( _
                        pUseEchoResultFormat, _
                        pProcessor, _
                        pBar, _
                        pBarnumber, _
                        pTimeframe, _
                        pNormaliseDailyBarTimestamps, _
                        pSecType, _
                        pTickSize, _
                        pIncludeMillisecs), _
                    True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gOutputBarToTextStream( _
                ByVal pProcessor As IProcessor, _
                ByVal pBar As Bar, _
                ByVal pTimeframe As TimePeriod, _
                ByVal pBarnumber As Long, _
                ByVal pNormaliseDailyBarTimestamps As Boolean, _
                ByVal pIncludeMillisecs As Boolean, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByVal pStream As TextStream)
Const ProcName As String = "gOutputBarToStdOut"
On Error GoTo Err

If pBar Is Nothing Then Exit Sub

pStream.WriteLine formatBar(False, pProcessor, pBar, pBarnumber, pTimeframe, pNormaliseDailyBarTimestamps, pSecType, pTickSize, pIncludeMillisecs)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gPerformFilenameVariableSubstitution( _
                ByVal pString As String, _
                ByVal pProcessor As IProcessor) As String
Const ProcName As String = "gPerformFilenameVariableSubstitution"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
Set lContractSpec = pProcessor.ContractSpec

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = SubstitutionVariablePattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pString)

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
    Case TodayVariable
        r = FormatTimestamp(todayDate, mTimestampDateOnlyFormat)
    Case YesterdayVariable
        r = FormatTimestamp(yesterdayDate, mTimestampDateOnlyFormat)
    Case FromDateVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampDateOnlyFormat)
    Case FromDateTimeVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampFormat + TimestampNoMillisecs)
    Case FromTimeVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampTimeOnlyFormat + TimestampNoMillisecs)
    Case ToDateVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampDateOnlyFormat)
        End If
    Case ToDateTimeVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampFormat + TimestampNoMillisecs)
        End If
    Case ToTimeVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampTimeOnlyFormat + TimestampNoMillisecs)
        End If
    Case TimeframeVariable
        r = pProcessor.Timeframe.ToShortString
    Case NumberRequestedVariable
        r = pProcessor.NumberOfBarsRequested
    Case NumberReturnedVariable
        r = pProcessor.NumberOfBarsOutput
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
                ByVal pProcessor As IProcessor, _
                ByVal pString As String, _
                ByVal pBar As Bar, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pBarnumber As Long, _
                ByVal pNormaliseDailyBarTimestamps As Boolean, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByVal pIncludeMillisecs As Boolean) As String
Const ProcName As String = "gPerformResultVariableSubstitution"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
Set lContractSpec = pProcessor.ContractSpec

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = SubstitutionVariablePattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pString)

Dim s As String
Dim lCurrPosn As Long: lCurrPosn = 1

Dim lMatch As Match
For Each lMatch In lMatches
    s = s & Mid$(pString, lCurrPosn, lMatch.FirstIndex - lCurrPosn + 1)
    lCurrPosn = lMatch.FirstIndex + lMatch.Length + 1
    
    Dim r As String
    
    Dim lVariable As String: lVariable = UCase$(lMatch.SubMatches(0))
    Select Case lVariable
    Case BarNumberVariable
        r = CStr(pBarnumber)
    Case TimestampVariable
        If Not pNormaliseDailyBarTimestamps Then
            r = FormatTimestamp(pBar.TimeStamp, mTimestampFormat Or (Not pIncludeMillisecs And TimestampNoMillisecs))
        ElseIf pTimePeriod.Units = TimePeriodDay Or _
                pTimePeriod.Units = TimePeriodWeek Or _
                pTimePeriod.Units = TimePeriodMonth Or _
                pTimePeriod.Units = TimePeriodYear Then
            r = FormatTimestamp(pBar.TimeStamp, mTimestampDateOnlyFormat)
        Else
            r = FormatTimestamp(pBar.TimeStamp, mTimestampFormat Or (Not pIncludeMillisecs And TimestampNoMillisecs))
        End If
    Case OpenVariable
        r = FormatPrice(pBar.OpenValue, pSecType, pTickSize)
    Case HighVariable
        r = FormatPrice(pBar.HighValue, pSecType, pTickSize)
    Case LowVariable
        r = FormatPrice(pBar.LowValue, pSecType, pTickSize)
    Case CloseVariable
        r = FormatPrice(pBar.CloseValue, pSecType, pTickSize)
    Case VolumeVariable
        r = CStr(pBar.Volume)
    Case TickVolumeVariable
        r = CStr(pBar.TickVolume)
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
    Case FromDateVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampDateOnlyFormat)
    Case FromDateTimeVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampFormat + TimestampNoMillisecs)
    Case FromTimeVariable
        r = FormatTimestamp(pProcessor.FromDate, mTimestampTimeOnlyFormat + TimestampNoMillisecs)
    Case ToDateVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampDateOnlyFormat)
        End If
    Case ToDateTimeVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampFormat + TimestampNoMillisecs)
        End If
    Case ToTimeVariable
        If pProcessor.ToDate = MaxDate Then
            r = DateLatest
        Else
            r = FormatTimestamp(pProcessor.ToDate, mTimestampTimeOnlyFormat + TimestampNoMillisecs)
        End If
    Case TimeframeVariable
        r = pProcessor.Timeframe.ToShortString
    Case NumberRequestedVariable
        r = pProcessor.NumberOfBarsRequested
    Case NumberReturnedVariable
        r = pProcessor.NumberOfBarsOutput
    Case NewlineVariable
        r = vbCrLf
    Case Default
        Assert False, "Unexpected substitution variable: " & lVariable
    End Select
    s = s & r
Next

gPerformResultVariableSubstitution = s & Right$(pString, Len(pString) - lCurrPosn + 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gWriteErrorLine( _
                ByVal pMessage As String)
mErrorStringBuilder.Append pMessage
Dim s As String: s = "Line " & mLineNumber & ": " & mErrorStringBuilder.ToString
gCon.WriteErrorLine s
LogMessage "StdErr: " & s
mErrorStringBuilder.Clear
End Sub

Public Sub gWriteErrorString( _
                ByVal pMessage As String)
mErrorStringBuilder.Append pMessage
End Sub

Public Sub gWriteLineToConsole( _
                ByVal pMessage As String)
gCon.WriteLineToConsole pMessage
LogMessage pMessage
End Sub

Public Sub gWriteLineToStdOut( _
                ByVal pMessage As String, _
                Optional ByVal pDontLogit As Boolean)
Const ProcName As String = "gWriteLineToStdOut"

If pMessage = "" Then Exit Sub
If Not pDontLogit Then LogMessage "StdOut: " & pMessage
gCon.WriteLine pMessage
End Sub

Public Sub Main()
On Error GoTo Err

Set mErrorStringBuilder = CreateStringBuilder

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "gbd"
SetupDefaultLogging Command

Set gCon = GetConsole

logProgramId

mNumber = AllBars
mTo = MaxDate
mNormaliseDailyBarTimestamps = True

mTimestampFormat = TimestampDateAndTimeISO8601
mTimestampDateOnlyFormat = TimestampDateOnlyISO8601
mTimestampTimeOnlyFormat = TimestampTimeOnlyISO8601

mCommandSeparator = ";"

Set mClp = CreateCommandLineParser(Command)

Dim lLogApiMessages As ApiMessageLoggingOptions
Dim lLogRawApiMessages As ApiMessageLoggingOptions
Dim lLogApiMessageStats As Boolean
If Not validateApiMessageLogging( _
                mClp.switchValue(SwitchApiMessageLogging), _
                lLogApiMessages, _
                lLogRawApiMessages, _
                lLogApiMessageStats) Then
    gWriteLineToConsole "API message logging setting is invalid"
    Exit Sub
End If

If mClp.Switch(SwitchLogToConsole) Then
    gLogToConsole = True
    DefaultLogLevel = LogLevelHighDetail
End If

If mClp.Switch("?") Or _
    mClp.NumberOfSwitches = 0 _
Then
    showUsage
    TerminateTWUtilities
    Exit Sub
End If

setupSubstitutionVariables

If mClp.Switch(SwitchCommandSeparator) Then mCommandSeparator = mClp.switchValue(SwitchCommandSeparator)
Assert Len(mCommandSeparator) = 1, "The command separator must be a single character"

If mClp.Switch(SwitchOutputPath) Then processOutputPathCommand mClp.switchValue(SwitchOutputPath)

If mClp.Switch(SwitchFromDb) Then
    mDataSource = FromDb
    If setupDbProviders(mClp.switchValue(SwitchFromDb)) Then process
ElseIf mClp.Switch(SwitchFromFile) Then
    mDataSource = FromFile
    If setupFileProviders(mClp.switchValue(SwitchFromFile)) Then process
ElseIf mClp.Switch(SwitchFromTws) Then
    mDataSource = FromTws
    If setupTwsProviders( _
            mClp.switchValue(SwitchFromTws), _
            lLogApiMessages, _
            lLogRawApiMessages, _
            lLogApiMessageStats) Then process
Else
    showUsage
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then
    gWriteErrorLine Err.Description
    gWriteErrorLine "At:"
    gWriteErrorLine Err.Source
End If

TerminateTWUtilities
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

lRegExp.Pattern = SubstitutionVariablePattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pPath)

Dim lMatch As Match
For Each lMatch In lMatches
    Dim lVariable As String: lVariable = lMatch.SubMatches(0)
    If Not isValidSubstitutionVariable(lVariable, mFilenameSubstitutionVariables, mMaxFilenameVariablesIndex) Then
        gWriteErrorLine lVariable & " is not a valid substitution variable"
        isValidPath = False
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidSubstitutableString( _
                ByVal pInput As String, _
                ByRef pSubstitutionVariables() As String, _
                ByVal pMaxVariablesIndex As Long) As Boolean
Const ProcName As String = "isValidSubstitutableString"
On Error GoTo Err

isValidSubstitutableString = True

Dim lRegExp As RegExp: Set lRegExp = gRegExp
lRegExp.IgnoreCase = True

lRegExp.Pattern = SubstitutionVariablePattern
lRegExp.Global = True

Dim lMatches As MatchCollection
Set lMatches = lRegExp.Execute(pInput)

Dim lMatch As Match
For Each lMatch In lMatches
    Dim lVariable As String: lVariable = lMatch.SubMatches(0)
    If Not isValidSubstitutionVariable(lVariable, pSubstitutionVariables, pMaxVariablesIndex) Then
        gWriteErrorLine lVariable & " is not a valid substitution variable"
        isValidSubstitutableString = False
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidSubstitutionVariable( _
                ByVal pString As String, _
                ByRef pSubstitutionVariables() As String, _
                ByVal pMaxVariablesIndex As Long) As Boolean
isValidSubstitutionVariable = BinarySearchStrings( _
                                UCase$(pString), _
                                pSubstitutionVariables, _
                                0, _
                                pMaxVariablesIndex + 1) >= 0
End Function

Private Function formatBar( _
                ByVal pUseEchoResultFormat As Boolean, _
                ByVal pProcessor As IProcessor, _
                ByVal pBar As Bar, _
                ByVal pBarnumber As Long, _
                ByVal pTimePeriod As TimePeriod, _
                ByVal pNormaliseDailyBarTimestamps As Boolean, _
                ByVal pSecType As SecurityTypes, _
                ByVal pTickSize As Double, _
                ByVal pIncludeMillisecs As Boolean) As String
Const ProcName As String = "formatBar"
On Error GoTo Err

If pBar Is Nothing Then Exit Function

Dim lResultFormat As String
lResultFormat = IIf(pUseEchoResultFormat, mStartEchoResultFormat, mStartResultFormat)
formatBar = gPerformResultVariableSubstitution( _
                        pProcessor, _
                        lResultFormat, _
                        pBar, _
                        pTimePeriod, _
                        pBarnumber, _
                        pNormaliseDailyBarTimestamps, _
                        pSecType, _
                        pTickSize, _
                        pIncludeMillisecs)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getInputLineAndWait( _
                Optional ByVal pDontReadInput As Boolean = False, _
                Optional ByVal pWaitTimeMIllisecs As Long = 5, _
                Optional ByVal pPrompt As String = ":") As String
Const ProcName As String = "getInputLine"
On Error GoTo Err

Dim lWaitUntilTime As Double
lWaitUntilTime = GetTimestampUTC + pWaitTimeMIllisecs / (86400# * 1000#)

If Not pDontReadInput Then getInputLineAndWait = Trim$(gCon.ReadLine(pPrompt))

Do
    ' allow queued system messages to be handled
    Wait 5
Loop Until GetTimestampUTC >= lWaitUntilTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gWriteLineToConsole s
LogMessage "Arguments: " & Command

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub logTasks()
Dim s As String

s = "Runnable tasks------------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetRunnableTaskSummary
s = s & "Processed tasks-----------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetProcessedTaskSummary
s = s & "Restartable tasks---------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetRestartableTaskSummary
s = s & "Pending tasks-------------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetPendingTaskSummary
s = s & "Suspended tasks-----------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetSuspendedTaskSummary
gWriteLineToConsole s
End Sub

Private Sub process()
Const ProcName As String = "process"
On Error GoTo Err

Dim lContinue As Boolean
processCommandLineCommands lContinue

If lContinue Then
    processStdInComands
End If

Do While mProcessors.Count <> 0
    getInputLineAndWait pDontReadInput:=True, pWaitTimeMIllisecs:=10
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processCommand(ByVal pCommandString As String)
Const ProcName As String = "processCommand"
On Error GoTo Err

Dim lCommand As String
lCommand = UCase$(Split(pCommandString, " ")(0))

Dim params As String
params = Trim$(Right$(pCommandString, Len(pCommandString) - Len(lCommand)))

Select Case lCommand
Case "TASKS"
    logTasks
Case ContractCommand
    processContractCommand params
Case FromCommand
    processFromCommand params
Case ToCommand
    processToCommand params
Case StartCommand
    processStartCommand params
Case StopCommand
    processStopCommand params
Case NumberCommand
    processNumberCommand params
Case TimeframeCommand
    processTimeframeCommand params
Case SessCommand
    processSessCommand
Case NonSessCommand
    processNonSessCommand
Case SessionOnlyCommand
    processSessionOnlyCommand params
Case MillisecsCommand
    processMillsecsCommand params
Case NoMillisecsCommand
    mIncludeMillisecs = False
Case HelpCommand, Help1Command
    showStdInHelp
Case SessionEndTimeCommand
    processSessionEndTimeCommand params
Case SessionStartTimeCommand
    processSessionStartTimeCommand params
Case DateOnlyCommmand
    processDateOnlyCommand params
Case OutpuPathCommand
    processOutputPathCommand params
Case AsyncCommand
    processAsyncCommand params
Case EntireSessionCommand
    processEntireSessionCommand params
Case DateTimeFormatCommand
    processDateTimeFormatCommand params
Case InFileCommand
    processInfileCommand params
Case EchoCommand
    processEchoCommand params
Case StartEchoCommand
    processStartEchoCommand params
Case ResultFormatCommand
    processResultFormatCommand params
Case EchoResultFormatCommand
    processEchoResultFormatCommand params
Case Else
    gWriteErrorLine "Invalid command '" & lCommand & "'"
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

Private Sub processContractCommand( _
                ByVal params As String)
Const ProcName As String = "processContractCommand"
On Error GoTo Err

If Trim$(params) = "" Then
    showContractHelp
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(params, ",")

If lClp.Arg(1) = "?" Or _
    lClp.Switch("?") Or _
    (lClp.NumberOfArgs = 0 And lClp.NumberOfSwitches = 0) _
Then
    showContractHelp
    Exit Sub
End If

If TryCreateContractSpecifierFromString(params, mContractSpec) Then

ElseIf lClp.NumberOfArgs > 1 Then
     Set mContractSpec = processPositionalContractString(lClp)
ElseIf lClp.NumberOfArgs = 1 Then
    Set mContractSpec = CreateContractSpecifierFromString(lClp.Arg(0))
Else
    Set lClp = CreateCommandLineParser(params, " ")
    If lClp.NumberOfSwitches = 0 Or _
        lClp.NumberOfArgs > 0 _
    Then
        gWriteErrorLine "Invalid contract syntax"
    Else
        Set mContractSpec = processTaggedContractString(lClp)
    End If
End If

Exit Sub

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Sub

Private Sub processDateOnlyCommand( _
                ByVal params As String)
params = UCase$(params)
If params = "" Or params = "YES" Or params = "TRUE" Or params = "ON" Then
    mNormaliseDailyBarTimestamps = True
ElseIf params = "NO" Or params = "FALSE" Or params = "OFF" Then
    mNormaliseDailyBarTimestamps = False
Else
    gWriteErrorLine "parameter must be YES, NO, ON, OFF, TRUE or FALSE"
End If
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

Private Sub processEchoCommand(ByVal params As String)
Const ProcName As String = "processEchoCommand"
On Error GoTo Err

gWriteLineToStdOut params

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processEchoResultFormatCommand(pParams As String)
If isValidSubstitutableString( _
                pParams, _
                mResultSubstitutionVariables, _
                mMaxResultVariablesIndex) Then mStartEchoResultFormat = pParams
End Sub

Private Sub processEntireSessionCommand(ByVal pParams As String)
pParams = UCase$(pParams)
If pParams = "" Or pParams = "YES" Or pParams = "TRUE" Or pParams = "ON" Then
    mEntireSession = True
ElseIf pParams = "NO" Or pParams = "FALSE" Or pParams = "OFF" Then
    mEntireSession = False
Else
    gWriteErrorLine "parameter must be YES, NO, ON, OFF, TRUE or FALSE"
End If
End Sub

Private Sub processFromCommand( _
                ByVal params As String)
Const ProcName As String = "processFromCommand"
On Error GoTo Err

params = UCase$(params)

mFrom = 0
mSpecialFrom = ""

If params = "" Then
ElseIf IsDate(params) Then
    mFrom = CDate(params)
ElseIf params = DateToday Then
    mSpecialFrom = params
ElseIf params = DateYesterday Then
    mSpecialFrom = params
ElseIf params = DateStartOfWeek Then
    mSpecialFrom = params
ElseIf params = DateStartOfPreviousWeek Then
    mSpecialFrom = params
Else
    gWriteErrorLine "Invalid from date '" & params & "'"
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

Private Sub processMillsecsCommand( _
                ByVal params As String)
params = UCase$(params)
If params = "" Or params = "YES" Or params = "TRUE" Or params = "ON" Then
    mIncludeMillisecs = True
ElseIf params = "NO" Or params = "FALSE" Or params = "OFF" Then
    mIncludeMillisecs = False
Else
    gWriteErrorLine "parameter must be YES, NO, ON, OFF, TRUE or FALSE"
End If
End Sub

Private Sub processNonSessCommand()
mSessionOnly = False
End Sub

Private Sub processNumberCommand( _
                ByVal params As String)
Const ProcName As String = "processNumberCommand"
On Error GoTo Err

If IsInteger(params, 1) Then
    mNumber = CLng(params)
ElseIf params = "-1" Or UCase$(params) = "ALL" Then
    mNumber = AllBars
Else
    gWriteErrorLine "Invalid number '" & params & "': must be an integer > 0 or -1 or 'ALL'"
End If

If mDataSource = FromFile Then gWriteLineToConsole "number command is ignored for tickfile input"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

Private Function processPositionalContractString( _
                ByVal pClp As CommandLineParser) As IContractSpecifier
Const ProcName As String = "processPositionalContractString"
On Error GoTo Err

'params: Shortname,Sectype,Exchange,Symbol,currency,Expiry,Multiplier,Strike,right,tradingclass

Dim lValidParams As Boolean: lValidParams = True

Dim lSectypeStr As String
lSectypeStr = Trim$(pClp.Arg(1))

Dim lExchange As String
lExchange = Trim$(pClp.Arg(2))

Dim lShortname As String
lShortname = Trim$(pClp.Arg(0))

Dim lSymbol As String
lSymbol = Trim$(pClp.Arg(3))

Dim lCurrencyCode As String
lCurrencyCode = Trim$(pClp.Arg(4))

Dim lExpiry As String
lExpiry = Trim$(pClp.Arg(5))

Dim lMultiplierStr As String
lMultiplierStr = Trim$(pClp.Arg(6))

Dim lStrikeStr As String
lStrikeStr = Trim$(pClp.Arg(7))

Dim lOptRightStr As String
lOptRightStr = Trim$(pClp.Arg(8))

Dim lTradingClass As String
lTradingClass = Trim$(pClp.Arg(9))

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gWriteErrorLine "Invalid lSectype '" & lSectypeStr & "'"
    lValidParams = False
End If

If lExpiry <> "" Then
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gWriteErrorLine "Invalid lExpiry '" & lExpiry & "'"
            lValidParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gWriteErrorLine "Invalid lExpiry '" & lExpiry & "'"
            lValidParams = False
        End If
    Else
        gWriteErrorLine "Invalid lExpiry '" & lExpiry & "'"
        lValidParams = False
    End If
End If
            
Dim lMultiplier As Double
If lMultiplierStr = "" Then
    lMultiplier = 0#
ElseIf IsNumeric(lMultiplierStr) Then
    lMultiplier = CDbl(lMultiplierStr)
Else
    gWriteErrorLine "Invalid lMultiplier '" & lMultiplierStr & "'"
    lValidParams = False
End If
            
Dim lStrike As Double
If lStrikeStr <> "" Then
    If IsNumeric(lStrikeStr) Then
        lStrike = CDbl(lStrikeStr)
    Else
        gWriteErrorLine "Invalid lStrike '" & lStrikeStr & "'"
        lValidParams = False
    End If
End If

Dim lOptRight As OptionRights
lOptRight = OptionRightFromString(lOptRightStr)
If lOptRightStr <> "" And lOptRight = OptNone Then
    gWriteErrorLine "Invalid right '" & lOptRightStr & "'"
    lValidParams = False
End If

        
If lValidParams Then
    Set processPositionalContractString = CreateContractSpecifier(lShortname, _
                                                                lSymbol, _
                                                                lTradingClass, _
                                                                lExchange, _
                                                                lSectype, _
                                                                lCurrencyCode, _
                                                                lExpiry, _
                                                                lMultiplier, _
                                                                lStrike, _
                                                                lOptRight)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processResultFormatCommand(pParams As String)
If isValidSubstitutableString( _
                pParams, _
                mResultSubstitutionVariables, _
                mMaxResultVariablesIndex) Then mStartResultFormat = pParams
End Sub

Private Sub processSessCommand()
mSessionOnly = True
End Sub

Private Sub processSessionOnlyCommand(ByVal pParams As String)
pParams = UCase$(pParams)
If pParams = "" Or pParams = "YES" Or pParams = "TRUE" Or pParams = "ON" Then
    mSessionOnly = True
ElseIf pParams = "NO" Or pParams = "FALSE" Or pParams = "OFF" Then
    mSessionOnly = False
Else
    gWriteErrorLine "parameter must be YES, NO, ON, OFF, TRUE or FALSE"
End If
End Sub

Private Sub processSessionEndTimeCommand(ByVal pParams As String)
Const ProcName As String = "processSessionEndTimeCommand"
On Error GoTo Err

If mDataSource = FromFile Then
    gWriteErrorLine "command ignored for this data source"
    Exit Sub
End If

If pParams = "" Then
    mSessionEndTime = 0
ElseIf IsDate(pParams) Then
    Dim lSessionTime: lSessionTime = CDate(pParams)
    If CDbl(lSessionTime) > Time235900 Or CDbl(lSessionTime) < 0# Then
        gWriteErrorLine "Invalid session end time '" & pParams & "': the value must be a time between 00:00 and 23:59"
    Else
        mSessionEndTime = lSessionTime
    End If
Else
    gWriteErrorLine "Invalid session end time '" & pParams & "' is not a date/time"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSessionStartTimeCommand(ByVal pParams As String)
Const ProcName As String = "processSessionStartTimeCommand"
On Error GoTo Err

If mDataSource = FromFile Then
    gWriteErrorLine "command ignored for this data source"
    Exit Sub
End If

If pParams = "" Then
    mSessionStartTime = 0
ElseIf IsDate(pParams) Then
    Dim lSessionTime: lSessionTime = CDate(pParams)
    If CDbl(lSessionTime) > Time235900 Or CDbl(lSessionTime) < 0# Then
        gWriteErrorLine "Invalid session start time '" & pParams & "': the value must be a time between 00:00 and 23:59"
    Else
        mSessionStartTime = lSessionTime
    End If
Else
    gWriteErrorLine "Invalid session start time '" & pParams & "' is not a date/time"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStartCommand( _
                ByVal params As String)
Const ProcName As String = "processStartCommand"
On Error GoTo Err

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(params, " ")

If mDataSource <> FromFile And mContractSpec Is Nothing Then
    gWriteErrorLine "Cannot start - no contract specified"
ElseIf mDataSource <> FromFile And mFrom = 0 And mSpecialFrom = "" And mNumber = AllBars Then
    gWriteErrorLine "Cannot start - either 'from' time or number of bars must be specified"
ElseIf mFrom > mTo And mTo <> 0 Then
    gWriteErrorLine "Cannot start - 'from' time must not be after 'to' time"
ElseIf mTimePeriod Is Nothing Then
    gWriteErrorLine "Cannot start - timeframe not specified"
ElseIf Not mAsync And mProcessors.Count <> 0 Then
    gWriteErrorLine "Cannot start - already running"
ElseIf lClp.NumberOfArgs > 2 Then
    gWriteErrorLine "Too many arguments"
Else
    
    Dim lPathAndFilename As String
    
    Dim lOutputToFiles As Boolean
    Dim lAppend As Boolean
    If lClp.NumberOfArgs = 0 Then
        lOutputToFiles = False
        lAppend = False
    ElseIf lClp.NumberOfArgs = 2 Then
        lPathAndFilename = lClp.Arg(1)
        If lClp.Arg(0) = AppendOperator Then
            lOutputToFiles = True
            lAppend = True
        ElseIf lClp.Arg(0) = OverwriteOperator Then
            lOutputToFiles = True
            lAppend = False
        Else
            gWriteErrorLine "First argument must be '>' or '>>'"
        End If
    ElseIf lClp.Arg(0) = AppendOperator Then
        lOutputToFiles = True
        lAppend = True
    ElseIf lClp.Arg(0) = OverwriteOperator Then
        lOutputToFiles = True
        lAppend = False
    Else
        lOutputToFiles = True
        lAppend = False
        lPathAndFilename = lClp.Arg(0)
    End If

    If isValidPath(lPathAndFilename) Then
        If mStartResultFormat = "" Then mStartResultFormat = DefaultResultFormat
        If mStartEchoResultFormat = "" Then mStartEchoResultFormat = DefaultResultFormat
        
        Dim lProcess As IProcessor
        If mDataSource = FromFile Then
            Dim lFileProcessor As New FileProcessor
            lFileProcessor.Initialise mTickfileName, mFrom, mTo, mNumber, mTimePeriod, mSessionOnly, mEntireSession, mIncludeMillisecs, mEchoToStdOut
            Set lProcess = lFileProcessor
        Else
            Dim lProcessor As New Processor
            lProcessor.Initialise mContractStore, _
                                mHistDataStore, _
                                mContractSpec, _
                                mFrom, _
                                mTo, _
                                mSpecialFrom, _
                                mSpecialTo, _
                                mNumber, _
                                mTimePeriod, _
                                mSessionOnly, _
                                mSessionStartTime, _
                                mSessionEndTime, _
                                mEntireSession, _
                                mNormaliseDailyBarTimestamps, _
                                mIncludeMillisecs, _
                                mEchoToStdOut
            Set lProcess = lProcessor
        End If
        
        mProcessors.Add lProcess
        lProcess.StartData IIf(lOutputToFiles, mOutputPath, ""), lPathAndFilename, lAppend
        If Not mAsync Then Set mCurrentProcessor = lProcess
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStartEchoCommand( _
                ByVal params As String)
mEchoToStdOut = True
processStartCommand params
End Sub

Private Sub processStdInComands()
Const ProcName As String = "processStdInComands"
On Error GoTo Err

Static sPreviousLineBlank As Boolean

Dim lInputString As String
lInputString = getInputLineAndWait(Not (mProviderReady And _
                                        (mAsync Or _
                                         mCurrentProcessor Is Nothing)))

Do Until lInputString = gCon.EofString Or _
        UCase$(lInputString) = ExitCommand
    
    If lInputString = "" Then
        If Not sPreviousLineBlank Then mLineNumber = mLineNumber + 1
        sPreviousLineBlank = True
            
        ' ignore blank lines, but echo them to StdOut when
        ' piping to another program
        If gCon.StdOutType = FileTypePipe Then gWriteLineToStdOut ""
    ElseIf Left$(lInputString, 1) = "#" Then
        sPreviousLineBlank = False
        mLineNumber = mLineNumber + 1
        LogMessage "con: " & lInputString
        ' ignore comments
    Else
        sPreviousLineBlank = False
        mLineNumber = mLineNumber + 1
        LogMessage "con: " & lInputString
        processCommand lInputString
    End If

    lInputString = getInputLineAndWait(Not (mProviderReady And _
                                            (mAsync Or _
                                             mCurrentProcessor Is Nothing)))
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStopCommand( _
                ByVal pParams As String)
Const ProcName As String = "processStopCommand"
On Error GoTo Err

If pParams = "" Then
    If Not mCurrentProcessor Is Nothing Then
        mCurrentProcessor.StopData
    ElseIf Not mAsync Then
        gWriteErrorLine "Error: nothing is running"
    Else
        gWriteErrorLine "Error: you must use 'STOP ALL' during async processing"
    End If
ElseIf UCase$(pParams) = "ALL" Then
    If mProcessors.Count = 0 Then
        gWriteErrorLine "Error: nothing is running"
    Else
        Dim lProcessor As IProcessor
        For Each lProcessor In mProcessors
            lProcessor.StopData
        Next
    End If
Else
    gWriteErrorLine "Error: the only parameter allowed is ALL"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processTaggedContractString( _
                ByVal pClp As CommandLineParser) As IContractSpecifier
Const ProcName As String = "processTaggedContractString"
On Error GoTo Err

Const CurrencySwitch                         As String = "CURRENCY"
Const CurrencySwitch1                        As String = "CURR"
Const ExchangeSwitch                         As String = "EXCHANGE"
Const ExchangeSwitch1                        As String = "EXCH"
Const ExpirySwitch                           As String = "EXPIRY"
Const ExpirySwitch1                          As String = "EXP"
Const LocalSymbolSwitch                      As String = "LOCALSYMBOL"
Const LocalSymbolSwitch1                     As String = "LOCAL"
Const MultiplierSwitch                       As String = "MULTIPLIER"
Const MultiplierSwitch1                      As String = "MULT"
Const RightSwitch                            As String = "RIGHT"
Const SecTypeSwitch                          As String = "SECTYPE"
Const SecTypeSwitch1                         As String = "SEC"
Const SymbolSwitch                           As String = "SYMBOL"
Const SymbolSwitch1                          As String = "SYMB"
Const StrikeSwitch                           As String = "STRIKE"
Const StrikeSwitch1                          As String = "STR"
Const TradingClassSwitch                     As String = "TRADINGCLASS"

If pClp.Arg(0) = "?" Or _
    pClp.Switch("?") Or _
    (pClp.Arg(0) = "" And pClp.NumberOfSwitches = 0) _
Then
    showContractHelp
    Exit Function
End If

Dim validParams As Boolean
validParams = True

Dim lSectypeStr As String: lSectypeStr = pClp.switchValue(SecTypeSwitch)
If lSectypeStr = "" Then lSectypeStr = pClp.switchValue(SecTypeSwitch1)

Dim lExchange As String: lExchange = pClp.switchValue(ExchangeSwitch)
If lExchange = "" Then lExchange = pClp.switchValue(ExchangeSwitch1)

Dim lLocalSymbol As String: lLocalSymbol = pClp.switchValue(LocalSymbolSwitch)
If lLocalSymbol = "" Then lLocalSymbol = pClp.switchValue(LocalSymbolSwitch1)

Dim lSymbol As String: lSymbol = pClp.switchValue(SymbolSwitch)
If lSymbol = "" Then lSymbol = pClp.switchValue(SymbolSwitch1)

Dim lCurrency As String: lCurrency = pClp.switchValue(CurrencySwitch)
If lCurrency = "" Then lCurrency = pClp.switchValue(CurrencySwitch1)

Dim lExpiry As String: lExpiry = pClp.switchValue(ExpirySwitch)
If lExpiry = "" Then lExpiry = pClp.switchValue(ExpirySwitch1)

Dim lMultiplier As String: lMultiplier = pClp.switchValue(MultiplierSwitch)
If lMultiplier = "" Then lMultiplier = pClp.switchValue(MultiplierSwitch1)
If lMultiplier = "" Then lMultiplier = "1.0"

Dim lStrike As String: lStrike = pClp.switchValue(StrikeSwitch)
If lStrike = "" Then lStrike = pClp.switchValue(StrikeSwitch1)
If lStrike = "" Then lStrike = "0.0"

Dim lRight As String: lRight = pClp.switchValue(RightSwitch)

Dim lTradingClass As String: lTradingClass = pClp.switchValue(TradingClassSwitch)

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gWriteErrorLine "Line " & mLineNumber & "Invalid Sectype '" & lSectypeStr & "'"
    validParams = False
End If

If lExpiry <> "" Then
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gWriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gWriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    Else
        gWriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
        validParams = False
    End If
End If
            
Dim Multiplier As Double
If lMultiplier = "" Then
    Multiplier = 0#
ElseIf IsNumeric(lMultiplier) Then
    Multiplier = CDbl(lMultiplier)
Else
    gWriteErrorLine "Line " & mLineNumber & "Invalid multiplier '" & lMultiplier & "'"
    validParams = False
End If
            
Dim Strike As Double
If lStrike <> "" Then
    If IsNumeric(lStrike) Then
        Strike = CDbl(lStrike)
    Else
        gWriteErrorLine "Line " & mLineNumber & "Invalid strike '" & lStrike & "'"
        validParams = False
    End If
End If

Dim optRight As OptionRights
optRight = OptionRightFromString(lRight)
If lRight <> "" And optRight = OptNone Then
    gWriteErrorLine "Line " & mLineNumber & "Invalid right '" & lRight & "'"
    validParams = False
End If

        
If validParams Then
    Set processTaggedContractString = CreateContractSpecifier(lLocalSymbol, _
                                                            lSymbol, _
                                                            lTradingClass, _
                                                            lExchange, _
                                                            lSectype, _
                                                            lCurrency, _
                                                            lExpiry, _
                                                            Multiplier, _
                                                            Strike, _
                                                            optRight)
End If

Exit Function

Err:
If Err.Number = ErrorCodes.ErrIllegalArgumentException Then
    gWriteErrorLine "Line " & mLineNumber & Err.Description
Else
    gHandleUnexpectedError ProcName, ModuleName
End If
End Function

Private Sub processTimeframeCommand( _
                ByVal params As String)
Const ProcName As String = "processTimeframeCommand"
On Error GoTo Err

If Trim$(params) = "" Then
    showTimeframeHelp
    Exit Sub
End If

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(params, " ")

If lClp.NumberOfArgs < 1 Then
    gWriteErrorLine "Invalid timeframe - the bar length must be supplied"
    Exit Sub
End If

If Not IsInteger(lClp.Arg(0), 1) Then
    gWriteErrorLine "Invalid bar length '" & Trim$(lClp.Arg(0)) & "': must be an integer > 0"
    Exit Sub
End If
Dim lBarLength As Long
lBarLength = CLng(lClp.Arg(0))

Dim lBarUnits As TimePeriodUnits
lBarUnits = TimePeriodMinute
If Trim$(lClp.Arg(1)) <> "" Then
    lBarUnits = TimePeriodUnitsFromString(lClp.Arg(1))
    If lBarUnits = TimePeriodNone Then
        gWriteErrorLine "Invalid bar units '" & Trim$(lClp.Arg(1)) & "': must be one of s,m,h,d,w,mm,v,tv,tm"
    Exit Sub
    End If
End If

Set mTimePeriod = GetTimePeriod(lBarLength, lBarUnits)

If mDataSource <> FromFile Then
    If Not mHistDataStore.TimePeriodValidator.IsValidTimePeriod(mTimePeriod) Then
        gWriteErrorLine ("Unsupported time period: " & mTimePeriod.ToString)
        Exit Sub
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub processToCommand( _
                ByVal params As String)
Const ProcName As String = "processToCommand"
On Error GoTo Err

params = UCase$(params)

mTo = 0
mSpecialTo = ""

If params = "" Then
ElseIf params = DateLatest Then
    mTo = MaxDate
ElseIf IsDate(params) Then
    mTo = CDate(params)
ElseIf params = DateToday Then
    mSpecialTo = params
ElseIf params = DateYesterday Then
    mSpecialTo = params
ElseIf params = DateTomorrow Then
    mSpecialTo = params
ElseIf params = DateEndOfWeek Then
    mSpecialTo = params
Else
    gWriteErrorLine "Invalid to date '" & params & "'"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupDbProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupDbProviders"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(switchValue, ",")

On Error Resume Next

Dim server As String
server = lClp.Arg(0)

Dim dbtypeStr As String
dbtypeStr = lClp.Arg(1)

Dim database As String
database = lClp.Arg(2)

Dim username As String
username = lClp.Arg(3)

Dim password As String
password = lClp.Arg(4)

On Error GoTo 0

If username <> "" And password = "" Then
    password = gCon.ReadLineFromConsole("Password:", "*")
End If
    
Dim dbtype As DatabaseTypes
dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gWriteErrorLine "Error: invalid dbtype"
    Exit Function
End If

On Error Resume Next

Dim lDbClient As DBClient
Set lDbClient = CreateTradingDBClient(dbtype, server, database, username, password, True)


Set mHistDataStore = lDbClient.HistoricalDataStore
Set mContractStore = lDbClient.ContractStore

setupDbProviders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupFileProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupFileProviders"
On Error GoTo Err

mTickfileName = switchValue
setupFileProviders = True
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setupSubstitutionVariables()
ReDim mFilenameSubstitutionVariables(15) As String
mMaxFilenameVariablesIndex = -1

addFilenameSubstitutionVariable ContractVariable
addFilenameSubstitutionVariable SymbolVariable
addFilenameSubstitutionVariable LocalSymbolVariable
addFilenameSubstitutionVariable SecTypeVariable
addFilenameSubstitutionVariable SecTypeAbbrvVariable
addFilenameSubstitutionVariable ExchangeVariable
addFilenameSubstitutionVariable ExpiryVariable
addFilenameSubstitutionVariable CurrencyVariable
addFilenameSubstitutionVariable MultiplierVariable
addFilenameSubstitutionVariable StrikeVariable
addFilenameSubstitutionVariable RightVariable
addFilenameSubstitutionVariable TodayVariable
addFilenameSubstitutionVariable YesterdayVariable
addFilenameSubstitutionVariable FromDateVariable
addFilenameSubstitutionVariable FromTimeVariable
addFilenameSubstitutionVariable FromDateTimeVariable
addFilenameSubstitutionVariable ToDateVariable
addFilenameSubstitutionVariable ToTimeVariable
addFilenameSubstitutionVariable ToDateTimeVariable
addFilenameSubstitutionVariable TimeframeVariable
addFilenameSubstitutionVariable NumberRequestedVariable
addFilenameSubstitutionVariable NumberReturnedVariable

SortStrings mFilenameSubstitutionVariables, EndIndex:=mMaxFilenameVariablesIndex


ReDim mResultSubstitutionVariables(15) As String
mMaxResultVariablesIndex = -1

addResultSubstitutionVariable BarNumberVariable
addResultSubstitutionVariable TimestampVariable
addResultSubstitutionVariable OpenVariable
addResultSubstitutionVariable HighVariable
addResultSubstitutionVariable LowVariable
addResultSubstitutionVariable CloseVariable
addResultSubstitutionVariable VolumeVariable
addResultSubstitutionVariable TickVolumeVariable
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
addResultSubstitutionVariable FromDateVariable
addResultSubstitutionVariable FromDateTimeVariable
addResultSubstitutionVariable FromTimeVariable
addResultSubstitutionVariable ToDateVariable
addResultSubstitutionVariable ToDateTimeVariable
addResultSubstitutionVariable ToTimeVariable
addResultSubstitutionVariable TimeframeVariable
addResultSubstitutionVariable NumberRequestedVariable
addResultSubstitutionVariable NumberReturnedVariable
addResultSubstitutionVariable NewlineVariable

SortStrings mResultSubstitutionVariables, EndIndex:=mMaxResultVariablesIndex
End Sub

Private Function setupTwsProviders( _
                ByVal switchValue As String, _
                ByVal pLogApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogRawApiMessages As ApiMessageLoggingOptions, _
                ByVal pLogApiMessageStats As Boolean) As Boolean
Const ProcName As String = "setupTwsProviders"
On Error GoTo Err

On Error Resume Next

setupTwsProviders = True

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(switchValue, ",")

Dim server As String
server = lClp.Arg(0)

Dim port As String
port = lClp.Arg(1)

Dim clientId As String
clientId = lClp.Arg(2)

Dim connectionRetryInterval As String
connectionRetryInterval = lClp.Arg(3)

On Error GoTo Err

If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gWriteErrorLine "Error: port must be an integer > 0"
    setupTwsProviders = False
End If
    
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gWriteErrorLine "Error: clientId must be an integer >= 0 and <= 999999999"
    setupTwsProviders = False
End If

If connectionRetryInterval = "" Then
ElseIf Not IsInteger(connectionRetryInterval, 0, 3600) Then
    gWriteErrorLine "Error: connection retry interval must be an integer >= 0 and <= 3600"
    setupTwsProviders = False
End If

If Not setupTwsProviders Then Exit Function

Dim lTwsClient As Client
If connectionRetryInterval = "" Then
    Set lTwsClient = GetClient( _
                            server, _
                            CLng(port), _
                            CLng(clientId), _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats)
Else
    Set lTwsClient = GetClient( _
                            server, _
                            CLng(port), _
                            CLng(clientId), _
                            pConnectionRetryIntervalSecs:=CLng(connectionRetryInterval), _
                            pLogApiMessages:=pLogApiMessages, _
                            pLogRawApiMessages:=pLogRawApiMessages, _
                            pLogApiMessageStats:=pLogApiMessageStats)
End If
Set mTWSConnectionMonitor = New TWSConnectionMonitor
lTwsClient.AddTwsConnectionStateListener mTWSConnectionMonitor

Set mHistDataStore = lTwsClient.GetHistoricalDataStore
'lTwsClient.DisableHistoricalDataRequestPacing

Set mContractStore = lTwsClient.GetContractStore
    
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showContractHelp()
gWriteLineToConsole "contract localsymbol[@exchange]"
gWriteLineToConsole "OR   "
gWriteLineToConsole "contract localsymbol@SMART/primaryexchange"
gWriteLineToConsole "OR   "
gWriteLineToConsole "contract localsymbol@<SMART|SMARTAUS|SMARTCAN|SMARTEUR|SMARTNASDAQ|SMARTNYSE|"
gWriteLineToConsole "                      SMARTUK|SMARTUS>"
gWriteLineToConsole "OR   "
gWriteLineToConsole "contract /specifier [/specifier]..."
gWriteLineToConsole "    where:"
gWriteLineToConsole "    specifier ::=   local[symbol]:STRING"
gWriteLineToConsole "                  | symb[ol]:STRING"
gWriteLineToConsole "                  | sec[type]:<STK|FUT|FOP|CASH|OPT>"
gWriteLineToConsole "                  | exch[ange]:STRING"
gWriteLineToConsole "                  | curr[ency]:<USD|EUR|GBP|JPY|CHF | etc>"
gWriteLineToConsole "                  | exp[iry]:<yyyymm|yyyymmdd|expiryoffset>"
gWriteLineToConsole "                  | mult[iplier]:INTEGER"
gWriteLineToConsole "                  | str[ike]:DOUBLE"
gWriteLineToConsole "                  | right:<CALL|PUT> "
gWriteLineToConsole "    expiryoffset ::= INTEGER(0..10)"
gWriteLineToConsole "OR   "
gWriteLineToConsole "contract localsymbol,sectype,exchange,symbol,currency,expiry,multiplier,strike,"
gWriteLineToConsole "         right"
gWriteLineToConsole ""
gWriteLineToConsole "Examples   "
gWriteLineToConsole "    contract ESH0"
gWriteLineToConsole "    contract FDAX MAR 20@DTB"
gWriteLineToConsole "    contract MSFT@SMARTUS"
gWriteLineToConsole "    contract MSFT@SMART/ISLAND"
gWriteLineToConsole "    contract /SYMBOL:MSFT /SECTYPE:OPT /EXCHANGE:CBOE /EXPIRY:20200117 "
gWriteLineToConsole "             /STRIKE:150 /RIGHT:C"
gWriteLineToConsole "    contract /SYMBOL:ES /SECTYPE:FUT /EXCHANGE:GLOBEX /EXPIRY:1 "
gWriteLineToConsole "    contract ,FUT,GLOBEX,ES,,1"

End Sub

Private Sub showStdInHelp()
gWriteLineToConsole "StdIn Format:"
gWriteLineToConsole ""
gWriteLineToConsole "#comment"

showContractHelp

gWriteLineToConsole "from starttime"
gWriteLineToConsole "to [endtime]"
gWriteLineToConsole "to LATEST"
gWriteLineToConsole "number n               # -1 or ALL => return all available bars"

showTimeframeHelp

gWriteLineToConsole "nonsess                # include bars outside session"
gWriteLineToConsole "sess                   # include only bars within the session"
gWriteLineToConsole "sessionstarttime time  # time of day the session is deemed to start:"
gWriteLineToConsole "                       # must between 00:00 and 23:59"
gWriteLineToConsole "sessionendtime time    # time of day the session is deemed to end:"
gWriteLineToConsole "                       # must between 00:00 and 23:59"
gWriteLineToConsole "millisecs              # include millisecs in bar timestamps"
gWriteLineToConsole "nomillisecs            # exclude millisecs in bar timestamps (default)"
gWriteLineToConsole "start"
gWriteLineToConsole "stop"
gWriteLineToConsole ""
gWriteLineToConsole "Note that if data is from TWS and sessionstarttime and/or"
gWriteLineToConsole "sessionendtime are not supplied, then the session times will be"
gWriteLineToConsole "deduced from IB's contract data, but ONLY if the contract has not"
gWriteLineToConsole "expired (IB does not supply this information for expired contracts)."
gWriteLineToConsole "Otherwise, the session is assumed to run from midnight to midnight."
gWriteLineToConsole "Since stock and index contracts don't expire, IB's session times "
gWriteLineToConsole "always apply unless overridden."
gWriteLineToConsole ""
gWriteLineToConsole "If data is from the TradeBuild historical database and sessionstarttime"
gWriteLineToConsole "and/or sessionendtime are not supplied, then the session times will be"
gWriteLineToConsole "as defined for the relevant contract in the TradeBuild contracts"
gWriteLineToConsole "database"
End Sub

Private Sub showTimeframeHelp()
gWriteLineToConsole "timeframe timeframespec"
gWriteLineToConsole "  where"
gWriteLineToConsole "    timeframespec  ::= length [units]"
gWriteLineToConsole "    units          ::=     s   seconds"
gWriteLineToConsole "                           m   minutes (default)"
gWriteLineToConsole "                           h   hours"
gWriteLineToConsole "                           d   days"
gWriteLineToConsole "                           w   weeks"
gWriteLineToConsole "                           mm   months"
gWriteLineToConsole "                           v   volume (constant volume bars)"
gWriteLineToConsole "                           tv  tick volume (constant tick volume bars)"
gWriteLineToConsole "                           tm   ticks movement (constant range bars)"
End Sub

Private Sub showUsage()
gWriteLineToConsole "Usage:"
gWriteLineToConsole "gbd27 -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gWriteLineToConsole "    OR"
gWriteLineToConsole "    -fromfile:tickfilepath"
gWriteLineToConsole "    OR"
gWriteLineToConsole "    -fromtws:[twsserver][,[port][,[clientid]]]"
gWriteLineToConsole ""
showStdInHelp
gWriteLineToConsole ""
gWriteLineToConsole "StdOut Format:"
gWriteLineToConsole ""
gWriteLineToConsole "timestamp,open,high,low,close,volume,tickvolume"
gWriteLineToConsole ""
gWriteLineToConsole "  where"
gWriteLineToConsole ""
gWriteLineToConsole "    timestamp ::= yyyy-mm-dd hh:mm:ss[.nnn]"
gWriteLineToConsole ""
gWriteLineToConsole ""
End Sub

Private Function todayDate() As Date
todayDate = Int(WorkingDayDate(WorkingDayNumber(Now), Now))
End Function

Private Function tomorrowDate() As Date
tomorrowDate = Int(WorkingDayDate(WorkingDayNumber(Now) + 1, Now))
End Function

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

Private Function yesterdayDate() As Date
yesterdayDate = Int(WorkingDayDate(WorkingDayNumber(Now) - 1, Now))
End Function





