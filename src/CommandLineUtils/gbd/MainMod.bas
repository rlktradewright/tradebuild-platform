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

Public Const ProjectName                    As String = "gbd"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const ContractCommand               As String = "CONTRACT"
Private Const FromCommand                   As String = "FROM"
Private Const ToCommand                     As String = "TO"
Private Const StartCommand                  As String = "START"
Private Const StopCommand                   As String = "STOP"
Private Const NumberCommand                 As String = "NUMBER"
Private Const TimeframeCommand              As String = "TIMEFRAME"
Private Const SessCommand                   As String = "SESS"
Private Const NonSessCommand                As String = "NONSESS"
Private Const MillisecsCommand              As String = "MILLISECS"
Private Const NoMillisecsCommand            As String = "NOMILLISECS"
Private Const HelpCommand                   As String = "HELP"
Private Const Help1Command                  As String = "?"

Private Const SwitchFromDb                  As String = "fromdb"
Private Const SwitchFromFile                As String = "fromfile"
Private Const SwitchFromTws                 As String = "fromtws"
Private Const SwitchLogToConsole            As String = "logtoconsole"

Private Const DefaultClientId               As Long = 205644991

'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev                                    As Boolean

Public gCon                                         As Console

Public gSecType                                     As SecurityTypes
Public gTickSize                                    As Double

Public gLogToConsole                                As Boolean

Public gProcessor                                   As IProcessor

Private mSwitch                                     As Switches

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

Public Sub gLogCompletion()
Const ProcName As String = "gLogCompletion"
On Error GoTo Err

gCon.WriteLineToConsole "Completed. Number of bars output = " & CStr(mNumberOfBarsWritten)

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
    mSwitch = FromDb
    If setupDbProviders(clp.switchValue(SwitchFromDb)) Then process
ElseIf clp.Switch(SwitchFromFile) Then
    mSwitch = FromFile
    If setupFileProviders(clp.switchValue(SwitchFromFile)) Then process
ElseIf clp.Switch(SwitchFromTws) Then
    mSwitch = FromTws
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

'params: shortname,sectype,exchange,symbol,currency,expiry,strike,right

If Trim$(params) = "" Then
    showContractHelp
    Exit Sub
End If

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(params, InputSep)

If clp.Arg(1) = "?" Or _
    clp.Switch("?") Or _
    clp.NumberOfArgs = 0 _
Then
    gCon.WriteLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
    Exit Sub
End If

Dim validParams As Boolean
validParams = True

Dim sectypeStr As String
sectypeStr = Trim$(clp.Arg(1))

Dim exchange As String
exchange = Trim$(clp.Arg(2))

Dim shortname As String
shortname = Trim$(clp.Arg(0))

Dim symbol As String
symbol = Trim$(clp.Arg(3))

Dim currencyCode As String
currencyCode = Trim$(clp.Arg(4))

Dim expiry As String
expiry = Trim$(clp.Arg(5))

Dim multiplierStr As String
multiplierStr = Trim$(clp.Arg(6))

Dim strikeStr As String
strikeStr = Trim$(clp.Arg(7))

Dim optRightStr As String
optRightStr = Trim$(clp.Arg(8))

Dim sectype As SecurityTypes
sectype = SecTypeFromString(sectypeStr)
If sectypeStr <> "" And sectype = SecTypeNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validParams = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiry = Format(CDate(expiry), "yyyymmdd")
    ElseIf Len(expiry) = 6 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Right$(expiry, 2) & "/01") Then
            gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
            validParams = False
        End If
    ElseIf Len(expiry) = 8 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
            validParams = False
        End If
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
        validParams = False
    End If
End If
            
Dim multiplier As Double
If multiplierStr = "" Then
    multiplier = 1#
ElseIf IsNumeric(multiplierStr) Then
    multiplier = CDbl(multiplierStr)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid multiplier '" & multiplierStr & "'"
    validParams = False
End If
            
Dim strike As Double
If strikeStr <> "" Then
    If IsNumeric(strikeStr) Then
        strike = CDbl(strikeStr)
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid strike '" & strikeStr & "'"
        validParams = False
    End If
End If

Dim optRight As OptionRights
optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid right '" & optRightStr & "'"
    validParams = False
End If

        
If validParams Then
    Set mContractSpec = CreateContractSpecifier(shortname, _
                                            symbol, _
                                            exchange, _
                                            sectype, _
                                            currencyCode, _
                                            expiry, _
                                            multiplier, _
                                            strike, _
                                            optRight)
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

If mSwitch = FromFile Then gCon.WriteLineToConsole "number command is ignored for tickfile input"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processSessCommand()
mSessionOnly = True
End Sub

Private Sub processStartCommand()
Const ProcName As String = "processStartCommand"
On Error GoTo Err

If mSwitch <> FromFile And mContractSpec Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mSwitch <> FromFile And mFrom = 0 And mNumber = 0 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - either 'from' time or number of bars must be specified"
ElseIf mFrom > mTo And mTo <> 0 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - 'from' time must not be after 'to' time"
ElseIf mTimePeriod Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - timeframe not specified"
ElseIf Not gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - already running"
Else
    mNumberOfBarsWritten = 0
       
    If mSwitch = FromFile Then
        Dim lFileProcessor As New FileProcessor
        lFileProcessor.Initialise mTickfileName, mFrom, mTo, mNumber, mTimePeriod, mSessionOnly
        Set gProcessor = lFileProcessor
    Else
        Dim lProcessor As New Processor
        lProcessor.Initialise mContractStore, mHistDataStore, mContractSpec, mFrom, mTo, mNumber, mTimePeriod, mSessionOnly
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

If mSwitch <> FromFile Then
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
gCon.WriteLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,multiplier,strike,right"
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
gCon.WriteLineToConsole "millisecs              # include millisecs in bar timestamps"
gCon.WriteLineToConsole "nomillisecs            # exclude millisecs in bar timestamps (default)"
gCon.WriteLineToConsole "start"
gCon.WriteLineToConsole "stop"
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
gCon.WriteLineToConsole "    -fromtws: [twsserver] [,[port][,[clientid]]]"
gCon.WriteLineToConsole ""
showStdInHelp
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "StdOut Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "timestamp,open,high,low,close,volume,tickvolume"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole ""
End Sub






