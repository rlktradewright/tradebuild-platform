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

Public Enum Switches
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
Private Const MillisecsCommand                      As String = "MILLISECS"
Private Const NoMillisecsCommand                    As String = "NOMILLISECS"
Private Const HelpCommand                           As String = "HELP"
Private Const Help1Command                          As String = "?"
Private Const SessionEndTimeCommand                 As String = "SESSIONENDTIME"
Private Const SessionStartTimeCommand               As String = "SESSIONSTARTTIME"

Private Const SwitchFromDb                          As String = "fromdb"
Private Const SwitchFromFile                        As String = "fromfile"
Private Const SwitchFromTws                         As String = "fromtws"
Private Const SwitchLogToConsole                    As String = "logtoconsole"

Private Const DefaultClientId                       As Long = 205644991

Private Const Time235900                            As Double = 0.99930556712963

'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev                                    As Boolean

Public gCon                                         As Console

Public gSecType                                     As SecurityTypes
Public gTickSize                                    As Double

Public gLogToConsole                                As Boolean

Public gProcessor                                   As IProcessor

Private mDataSource                                 As Switches

Private mTickfileName                               As String

Private mLineNumber                                 As Long
Private mContractSpec                               As IContractSpecifier
Private mFrom                                       As Date
Private mTo                                         As Date
Private mNumber                                     As Long
Private mTimePeriod                                 As TimePeriod
Private mSessionOnly                                As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mIncludeMillisecs                           As Boolean

Private mHistDataStore                              As IHistoricalDataStore
Private mContractStore                              As IContractStore

Private mNumberOfBarsWritten                        As Long

Private mSessionEndTime                             As Date
Private mSessionStartTime                           As Date

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

Public Function gGetContractName(ByVal pContractSpec As IContractSpecifier) As String
AssertArgument Not pContractSpec Is Nothing
gGetContractName = pContractSpec.LocalSymbol & _
                    IIf(pContractSpec.Exchange <> "", "@" & pContractSpec.Exchange, "")
End Function

Public Sub gHandleFatalError(ev As ErrorEventData)
On Error Resume Next    ' ignore any further errors that might arise

gCon.WriteErrorString "Error "
gCon.WriteErrorString CStr(ev.ErrorCode)
gCon.WriteErrorString ": "
gCon.WriteErrorLine ev.ErrorMessage
gCon.WriteErrorLine "At:"
gCon.WriteErrorLine ev.ErrorSource

' kill off any timers
'TerminateTWUtilities

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

'If mIsInDev Then
'    End
'Else
'    EndProcess
'End If

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

Public Sub gLogCompletion(ByVal pContractSpec As IContractSpecifier)
Const ProcName As String = "gLogCompletion"
On Error GoTo Err

gCon.WriteLineToConsole "Fetch completed for contract: " & gGetContractName(pContractSpec)
gCon.WriteLineToConsole "Number of bars output:  " & CStr(mNumberOfBarsWritten)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gLogDataRetrieved()
Const ProcName As String = "gLogDataRetrieved"
On Error GoTo Err

gCon.WriteLineToConsole "Data retrieved from source"

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

Public Sub gOutputBar(ByVal pBar As Bar)
Const ProcName As String = "gOutputBar"
On Error GoTo Err

If pBar Is Nothing Then Exit Sub

gCon.WriteString FormatTimestamp(pBar.TimeStamp, TimestampDateAndTimeISO8601 Or (Not mIncludeMillisecs And TimestampNoMillisecs))
gCon.WriteString ","
gCon.WriteString FormatPrice(pBar.OpenValue, gSecType, gTickSize)
gCon.WriteString ","
gCon.WriteString FormatPrice(pBar.HighValue, gSecType, gTickSize)
gCon.WriteString ","
gCon.WriteString FormatPrice(pBar.LowValue, gSecType, gTickSize)
gCon.WriteString ","
gCon.WriteString FormatPrice(pBar.CloseValue, gSecType, gTickSize)
gCon.WriteString ","
gCon.WriteString pBar.Volume
gCon.WriteString ","
gCon.WriteString pBar.TickVolume
gCon.WriteString ","
gCon.WriteLine pBar.OpenInterest

mNumberOfBarsWritten = mNumberOfBarsWritten + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Main()
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)

InitialiseTWUtilities

Set mFatalErrorHandler = New FatalErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "gbd"
SetupDefaultLogging command

'EnableTracing "tradebuild"
'EnableTracing "tickfilesp"

mNumber = &H7FFFFFFF

Set gCon = GetConsole

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(command)

If clp.Switch(SwitchLogToConsole) Then
    gLogToConsole = True
    DefaultLogLevel = LogLevelHighDetail
End If

If clp.Switch("?") Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf clp.Switch(SwitchFromDb) Then
    mDataSource = FromDb
    If setupDbProviders(clp.switchValue(SwitchFromDb)) Then process
ElseIf clp.Switch(SwitchFromFile) Then
    mDataSource = FromFile
    If setupFileProviders(clp.switchValue(SwitchFromFile)) Then process
ElseIf clp.Switch(SwitchFromTws) Then
    mDataSource = FromTws
    If setupTwsProviders(clp.switchValue(SwitchFromTws)) Then process
Else
    showUsage
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then
    gCon.WriteErrorLine Err.Description
    gCon.WriteErrorLine "At:"
    gCon.WriteErrorLine Err.Source
End If

TerminateTWUtilities
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function

Private Sub process()
Dim inString As String
Dim command As String
Dim params As String

Const ProcName As String = "process"
On Error GoTo Err

inString = Trim$(gCon.ReadLine(":"))
Do While inString <> gCon.EofString
    mLineNumber = mLineNumber + 1
    If inString = "" Then
        ' ignore blank lines
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments
    Else
        ' process command
        command = UCase$(Split(inString, " ")(0))
        params = Trim$(Right$(inString, Len(inString) - Len(command)))
        Select Case command
        Case ContractCommand
            processContractCommand params
        Case FromCommand
            processFromCommand params
        Case ToCommand
            processToCommand params
        Case StartCommand
            processStartCommand
        Case StopCommand
            processStopCommand
        Case NumberCommand
            processNumberCommand params
        Case TimeframeCommand
            processTimeframeCommand params
        Case SessCommand
            processSessCommand
        Case NonSessCommand
            processNonSessCommand
        Case MillisecsCommand
            mIncludeMillisecs = True
        Case NoMillisecsCommand
            mIncludeMillisecs = False
        Case HelpCommand, Help1Command
            showStdInHelp
        Case SessionEndTimeCommand
            processSessionEndTimeCommand params
        Case SessionStartTimeCommand
            processSessionStartTimeCommand params
        Case Else
            gCon.WriteErrorLine "Invalid command '" & command & "'"
        End Select
    End If
    inString = Trim$(gCon.ReadLine(":"))
    Wait 10
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processContractCommand( _
                ByVal params As String)
Const ProcName As String = "processContractCommand"
On Error GoTo Err

If Trim$(params) = "" Then
    showContractHelp
    Exit Sub
End If

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(params, ",")

If clp.Arg(1) = "?" Or _
    clp.Switch("?") Or _
    clp.NumberOfArgs = 0 _
Then
    showContractHelp
    Exit Sub
End If

If clp.NumberOfArgs > 1 Then
     Set mContractSpec = processPositionalContractString(clp)
ElseIf clp.NumberOfArgs = 1 Then
    Set mContractSpec = CreateContractSpecifierFromString(clp.Arg(0))
Else
    Set clp = CreateCommandLineParser(params, " ")
    If clp.NumberOfSwitches = 0 Or _
        clp.NumberOfArgs > 0 _
    Then
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid contract syntax"
    Else
        Set mContractSpec = processTaggedContractString(clp)
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processFromCommand( _
                ByVal params As String)
Const ProcName As String = "processFromCommand"
On Error GoTo Err

If params = "" Then
    mFrom = 0
ElseIf IsDate(params) Then
    mFrom = CDate(params)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid from date '" & params & "'"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
    mNumber = &H7FFFFFFF
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid number '" & params & "'" & ": must be an integer > 0 or -1"
End If

If mDataSource = FromFile Then gCon.WriteLineToConsole "number command is ignored for tickfile input"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function processPositionalContractString( _
                ByVal pClp As CommandLineParser) As IContractSpecifier
Const ProcName As String = "processPositionalContractString"
On Error GoTo Err

'params: lShortname,lSectype,lExchange,lSymbol,currency,lExpiry,lMultiplier,lStrike,right

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

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lSectype '" & lSectypeStr & "'"
    lValidParams = False
End If

If lExpiry <> "" Then
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lExpiry '" & lExpiry & "'"
            lValidParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lExpiry '" & lExpiry & "'"
            lValidParams = False
        End If
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lExpiry '" & lExpiry & "'"
        lValidParams = False
    End If
End If
            
Dim lMultiplier As Double
If lMultiplierStr = "" Then
    lMultiplier = 1#
ElseIf IsNumeric(lMultiplierStr) Then
    lMultiplier = CDbl(lMultiplierStr)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lMultiplier '" & lMultiplierStr & "'"
    lValidParams = False
End If
            
Dim lStrike As Double
If lStrikeStr <> "" Then
    If IsNumeric(lStrikeStr) Then
        lStrike = CDbl(lStrikeStr)
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid lStrike '" & lStrikeStr & "'"
        lValidParams = False
    End If
End If

Dim lOptRight As OptionRights
lOptRight = OptionRightFromString(lOptRightStr)
If lOptRightStr <> "" And lOptRight = OptNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid right '" & lOptRightStr & "'"
    lValidParams = False
End If

        
If lValidParams Then
    Set processPositionalContractString = CreateContractSpecifier(lShortname, _
                                                                lSymbol, _
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

Private Sub processSessCommand()
mSessionOnly = True
End Sub

Private Sub processSessionEndTimeCommand(ByVal pParams As String)
Const ProcName As String = "processSessionEndTimeCommand"
On Error GoTo Err

If mDataSource = FromFile Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": command ignored for this data source"
    Exit Sub
End If

If pParams = "" Then
    mSessionEndTime = 0
ElseIf IsDate(pParams) Then
    Dim lSessionTime: lSessionTime = CDate(pParams)
    If CDbl(lSessionTime) > Time235900 Or CDbl(lSessionTime) < 0# Then
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid session start time '" & pParams & "': the value must be a time between 00:00 and 23:59"
    Else
        mSessionEndTime = lSessionTime
    End If
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid session start time '" & pParams & "' is not a date/time"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSessionStartTimeCommand(ByVal pParams As String)
Const ProcName As String = "processSessionStartTimeCommand"
On Error GoTo Err

If mDataSource = FromFile Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": command ignored for this data source"
    Exit Sub
End If

If pParams = "" Then
    mSessionStartTime = 0
ElseIf IsDate(pParams) Then
    Dim lSessionTime: lSessionTime = CDate(pParams)
    If CDbl(lSessionTime) > Time235900 Or CDbl(lSessionTime) < 0# Then
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid session start time '" & pParams & "': the value must be a time between 00:00 and 23:59"
    Else
        mSessionStartTime = lSessionTime
    End If
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid session start time '" & pParams & "' is not a date/time"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStartCommand()
Const ProcName As String = "processStartCommand"
On Error GoTo Err

If mDataSource <> FromFile And mContractSpec Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mDataSource <> FromFile And mFrom = 0 And mNumber = 0 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - either 'from' time or number of bars must be specified"
ElseIf mFrom > mTo And mTo <> 0 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - 'from' time must not be after 'to' time"
ElseIf mTimePeriod Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - timeframe not specified"
ElseIf Not gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - already running"
Else
    mNumberOfBarsWritten = 0
       
    If mDataSource = FromFile Then
        Dim lFileProcessor As New FileProcessor
        lFileProcessor.Initialise mTickfileName, mFrom, mTo, mNumber, mTimePeriod, mSessionOnly
        Set gProcessor = lFileProcessor
    Else
        Dim lProcessor As New Processor
        lProcessor.Initialise mContractStore, _
                            mHistDataStore, _
                            mContractSpec, _
                            mFrom, _
                            mTo, _
                            mNumber, _
                            mTimePeriod, _
                            mSessionOnly, _
                            mSessionStartTime, _
                            mSessionEndTime
        Set gProcessor = lProcessor
    End If
    
    gProcessor.StartData
    
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processStopCommand()
Const ProcName As String = "processStopCommand"
On Error GoTo Err

If gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot stop - not started"
Else
    gProcessor.StopData
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

Dim lSectype As SecurityTypes
lSectype = SecTypeFromString(lSectypeStr)
If lSectypeStr <> "" And lSectype = SecTypeNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & "Invalid Sectype '" & lSectypeStr & "'"
    validParams = False
End If

If lExpiry <> "" Then
    If IsValidExpiry(lExpiry) Then
    ElseIf IsDate(lExpiry) Then
        lExpiry = Format(CDate(lExpiry), "yyyymmdd")
    ElseIf Len(lExpiry) = 6 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Right$(lExpiry, 2) & "/01") Then
            gCon.WriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    ElseIf Len(lExpiry) = 8 Then
        If Not IsDate(Left$(lExpiry, 4) & "/" & Mid$(lExpiry, 5, 2) & "/" & Right$(lExpiry, 2)) Then
            gCon.WriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
            validParams = False
        End If
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & "Invalid Expiry '" & lExpiry & "'"
        validParams = False
    End If
End If
            
Dim Multiplier As Double
If lMultiplier = "" Then
    Multiplier = 1#
ElseIf IsNumeric(lMultiplier) Then
    Multiplier = CDbl(lMultiplier)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & "Invalid multiplier '" & lMultiplier & "'"
    validParams = False
End If
            
Dim Strike As Double
If lStrike <> "" Then
    If IsNumeric(lStrike) Then
        Strike = CDbl(lStrike)
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & "Invalid strike '" & lStrike & "'"
        validParams = False
    End If
End If

Dim optRight As OptionRights
optRight = OptionRightFromString(lRight)
If lRight <> "" And optRight = OptNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & "Invalid right '" & lRight & "'"
    validParams = False
End If

        
If validParams Then
    Set processTaggedContractString = CreateContractSpecifier(lLocalSymbol, _
                                                            lSymbol, _
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
    gCon.WriteErrorLine "Line " & mLineNumber & Err.Description
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

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(params, " ")

If clp.NumberOfArgs < 1 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid timeframe - the bar length must be supplied"
    Exit Sub
End If

If Not IsInteger(clp.Arg(0), 1) Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid bar length '" & Trim$(clp.Arg(0)) & "': must be an integer > 0"
    Exit Sub
End If
Dim lBarLength As Long
lBarLength = CLng(clp.Arg(0))

Dim lBarUnits As TimePeriodUnits
lBarUnits = TimePeriodMinute
If Trim$(clp.Arg(1)) <> "" Then
    lBarUnits = TimePeriodUnitsFromString(clp.Arg(1))
    If lBarUnits = TimePeriodNone Then
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid bar units '" & Trim$(clp.Arg(1)) & "': must be one of s,m,h,d,w,mm,v,tv,tm"
    Exit Sub
    End If
End If

Set mTimePeriod = GetTimePeriod(lBarLength, lBarUnits)

If mDataSource <> FromFile Then
    If Not mHistDataStore.TimePeriodValidator.IsValidTimePeriod(mTimePeriod) Then
        gCon.WriteErrorLine ("Unsupported time period: " & mTimePeriod.ToString)
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

If params = "" Then
    mTo = 0
ElseIf UCase$(params) = "LATEST" Then
    mTo = MaxDate
ElseIf IsDate(params) Then
    mTo = CDate(params)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid to date '" & params & "'"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupDbProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupDbProviders"
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

On Error Resume Next

Dim server As String
server = clp.Arg(0)

Dim dbtypeStr As String
dbtypeStr = clp.Arg(1)

Dim database As String
database = clp.Arg(2)

Dim username As String
username = clp.Arg(3)

Dim password As String
password = clp.Arg(4)

On Error GoTo 0

If username <> "" And password = "" Then
    password = gCon.ReadLineFromConsole("Password:", "*")
End If
    
Dim dbtype As DatabaseTypes
dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gCon.WriteErrorLine "Error: invalid dbtype"
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

Private Function setupTwsProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupTwsProviders"
On Error GoTo Err

On Error Resume Next

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

Dim server As String
server = clp.Arg(0)

Dim port As String
port = clp.Arg(1)

Dim clientId As String
clientId = clp.Arg(2)

On Error GoTo Err

If port = "" Then
    port = 7496
ElseIf Not IsInteger(port, 0) Then
    gCon.WriteErrorLine "Error: port must be an integer > 0"
    setupTwsProviders = False
End If
    
If clientId = "" Then
    clientId = DefaultClientId
ElseIf Not IsInteger(clientId, 0, 999999999) Then
    gCon.WriteErrorLine "Error: clientId must be an integer >= 0 and <= 999999999"
    setupTwsProviders = False
End If

Dim lTwsClient As Client
Set lTwsClient = GetClient(server, CLng(port), CLng(clientId))

Set mHistDataStore = lTwsClient.GetHistoricalDataStore
Set mContractStore = lTwsClient.GetContractStore
    
setupTwsProviders = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showContractHelp()
gCon.WriteLineToConsole "contract localsymbol[@exchange]"
gCon.WriteLineToConsole "OR   "
gCon.WriteLineToConsole "contract localsymbol@SMART/primaryexchange"
gCon.WriteLineToConsole "OR   "
gCon.WriteLineToConsole "contract localsymbol@<SMART|SMARTAUS|SMARTCAN|SMARTEUR|SMARTNASDAQ|SMARTNYSE|"
gCon.WriteLineToConsole "                      SMARTUK|SMARTUS>"
gCon.WriteLineToConsole "OR   "
gCon.WriteLineToConsole "contract /specifier [/specifier]..."
gCon.WriteLineToConsole "    where:"
gCon.WriteLineToConsole "    specifier ::=   local[symbol]:STRING"
gCon.WriteLineToConsole "                  | symb[ol]:STRING"
gCon.WriteLineToConsole "                  | sec[type]:<STK|FUT|FOP|CASH|OPT>"
gCon.WriteLineToConsole "                  | exch[ange]:STRING"
gCon.WriteLineToConsole "                  | curr[ency]:<USD|EUR|GBP|JPY|CHF | etc>"
gCon.WriteLineToConsole "                  | exp[iry]:<yyyymm|yyyymmdd|expiryoffset>"
gCon.WriteLineToConsole "                  | mult[iplier]:INTEGER"
gCon.WriteLineToConsole "                  | str[ike]:DOUBLE"
gCon.WriteLineToConsole "                  | right:<CALL|PUT> "
gCon.WriteLineToConsole "    expiryoffset ::= INTEGER(0..10)"
gCon.WriteLineToConsole "OR   "
gCon.WriteLineToConsole "contract localsymbol,sectype,exchange,symbol,currency,expiry,multiplier,strike,"
gCon.WriteLineToConsole "         right"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "Examples   "
gCon.WriteLineToConsole "    contract ESH0"
gCon.WriteLineToConsole "    contract FDAX MAR 20@DTB"
gCon.WriteLineToConsole "    contract MSFT@SMARTUS"
gCon.WriteLineToConsole "    contract MSFT@SMART/ISLAND"
gCon.WriteLineToConsole "    contract /SYMBOL:MSFT /SECTYPE:OPT /EXCHANGE:CBOE /EXPIRY:20200117 "
gCon.WriteLineToConsole "             /STRIKE:150 /RIGHT:C"
gCon.WriteLineToConsole "    contract /SYMBOL:ES /SECTYPE:FUT /EXCHANGE:GLOBEX /EXPIRY:1 "
gCon.WriteLineToConsole "    contract ,FUT,GLOBEX,ES,,1"

End Sub

Private Sub showStdInHelp()
gCon.WriteLineToConsole "StdIn Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "#comment"

showContractHelp

gCon.WriteLineToConsole "from starttime"
gCon.WriteLineToConsole "to [endtime]"
gCon.WriteLineToConsole "to LATEST"
gCon.WriteLineToConsole "number n               # -1 or ALL => return all available bars"

showTimeframeHelp

gCon.WriteLineToConsole "nonsess                # include bars outside session"
gCon.WriteLineToConsole "sess                   # include only bars within the session"
gCon.WriteLineToConsole "sessionstarttime time  # time of day the session is deemed to start:"
gCon.WriteLineToConsole "                       # must between 00:00 and 23:59"
gCon.WriteLineToConsole "sessionendtime time    # time of day the session is deemed to end:"
gCon.WriteLineToConsole "                       # must between 00:00 and 23:59"
gCon.WriteLineToConsole "millisecs              # include millisecs in bar timestamps"
gCon.WriteLineToConsole "nomillisecs            # exclude millisecs in bar timestamps (default)"
gCon.WriteLineToConsole "start"
gCon.WriteLineToConsole "stop"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "Note that if data is from TWS and sessionstarttime and/or"
gCon.WriteLineToConsole "sessionendtime are not supplied, then the session times will be"
gCon.WriteLineToConsole "deduced from IB's contract data, but ONLY if the contract has not"
gCon.WriteLineToConsole "expired (IB does not supply this information for expired contracts."
gCon.WriteLineToConsole "Otherwise, the session is assumed to run from midnight to midnight."
gCon.WriteLineToConsole "Since stock and index contracts don't expire, IB's session times "
gCon.WriteLineToConsole "always apply unless overridden."
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "If data is from the TradeBuild historical database and sessionstarttime"
gCon.WriteLineToConsole "and/or sessionendtime are not supplied, then the session times will be"
gCon.WriteLineToConsole "as defined for the relevant contract in the TradeBuild contracts"
gCon.WriteLineToConsole "database"
End Sub

Private Sub showTimeframeHelp()
gCon.WriteLineToConsole "timeframe timeframespec"
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole "    timeframespec  ::= length [units]"
gCon.WriteLineToConsole "    units          ::=     s   seconds"
gCon.WriteLineToConsole "                           m   minutes (default)"
gCon.WriteLineToConsole "                           h   hours"
gCon.WriteLineToConsole "                           d   days"
gCon.WriteLineToConsole "                           w   weeks"
gCon.WriteLineToConsole "                           mm   months"
gCon.WriteLineToConsole "                           v   volume (constant volume bars)"
gCon.WriteLineToConsole "                           tv  tick volume (constant tick volume bars)"
gCon.WriteLineToConsole "                           tm   ticks movement (constant range bars)"
End Sub

Private Sub showUsage()
gCon.WriteLineToConsole "Usage:"
gCon.WriteLineToConsole "gbd27 -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.WriteLineToConsole "    OR"
gCon.WriteLineToConsole "    -fromfile:tickfilepath"
gCon.WriteLineToConsole "    OR"
gCon.WriteLineToConsole "    -fromtws:[twsserver][,[port][,[clientid]]]"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "StdOut Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "timestamp,open,high,low,close,volume,tickvolume"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    timestamp ::= yyyy-mm-dd hh:mm:ss[.nnn]"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole ""
End Sub






