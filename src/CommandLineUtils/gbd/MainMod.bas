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
Private Const HelpCommand                   As String = "HELP"
Private Const Help1Command                  As String = "?"

Private Const SwitchFromDb                  As String = "fromdb"
Private Const SwitchFromFile                As String = "fromfile"
Private Const SwitchFromTws                 As String = "fromtws"
Private Const SwitchLogToConsole            As String = "logtoconsole"

'@================================================================================
' Member variables
'@================================================================================

Private mIsInDev                                    As Boolean

Public gCon                                         As Console

Private mSwitch As Switches

Private mTickfileName As String

Private mLineNumber As Long
Private mContractSpec As IContractSpecifier
Private mFrom As Date
Private mTo As Date
Private mNumber As Long
Private mBarLength As Long
Private mBarUnits As TimePeriodUnits
Private mSessionOnly As Boolean

' this is public so that the Processor object can
' kill itself when it has finished replaying
Public gProcessor As Processor

Public gLogToConsole As Boolean

Private mFatalErrorHandler                          As FatalErrorHandler

Private mTB                                         As TradeBuildAPI

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
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.number)

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
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
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

mTo = MaxDate
mNumber = -1

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
    If setupServiceProviders(clp.switchValue(SwitchFromDb)) Then process
ElseIf clp.Switch(SwitchFromFile) Then
    mSwitch = FromFile
    If setupServiceProviders(clp.switchValue(SwitchFromFile)) Then process
ElseIf clp.Switch(SwitchFromTws) Then
    mSwitch = FromTws
    If setupServiceProviders(clp.switchValue(SwitchFromTws)) Then process
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
'params: shortname,sectype,exchange,symbol,currency,expiry,strike,right
Dim validParams As Boolean
Dim clp As CommandLineParser
Dim shortname As String
Dim sectypeStr As String
Dim sectype As SecurityTypes
Dim exchange As String
Dim symbol As String
Dim currencyCode As String
Dim expiry As String
Dim strikeStr As String
Dim strike As Double
Dim optRightStr As String
Dim optRight As OptionRights

Const ProcName As String = "processContractCommand"
On Error GoTo Err

If Not gProcessor Is Nothing Then
    showContractHelp
    Exit Sub
End If

Set clp = CreateCommandLineParser(params, InputSep)

If clp.Arg(1) = "?" Or _
    clp.Switch("?") Or _
    clp.NumberOfArgs = 0 _
Then
    gCon.WriteLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
    Exit Sub
End If

validParams = True

sectypeStr = Trim$(clp.Arg(1))
exchange = Trim$(clp.Arg(2))
shortname = Trim$(clp.Arg(0))
symbol = Trim$(clp.Arg(3))
currencyCode = Trim$(clp.Arg(4))
expiry = Trim$(clp.Arg(5))
strikeStr = Trim$(clp.Arg(6))
optRightStr = Trim$(clp.Arg(7))

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
            
If strikeStr <> "" Then
    If IsNumeric(strikeStr) Then
        strike = CDbl(strikeStr)
    Else
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid strike '" & strikeStr & "'"
        validParams = False
    End If
End If

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
                                            strike, _
                                            optRight)
End If

'Exit Sub
'
'Err:
'Set mContractSpec = Nothing
'gCon.WriteErrorLine "Error: " & Err.Description

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processFromCommand( _
                ByVal params As String)
Const ProcName As String = "processFromCommand"
On Error GoTo Err

If IsDate(params) Then
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

If Not IsInteger(params, 1) And params <> "-1" Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid number '" & params & "'" & ": must be an integer > 0"
Else
    mNumber = CLng(params)
    If mSwitch = FromFile Then gCon.WriteLineToConsole "number command is ignored for tickfile input"
End If

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
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - either from time or number of bars must be specified"
ElseIf mBarUnits = TimePeriodNone Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - timeframe not specified"
ElseIf Not gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - already running"
Else
       
    Set gProcessor = New Processor
    gProcessor.Initialise mTB
    
    Select Case mSwitch
    Case FromDb, FromTws
        Dim sbp As New StreamBasedProcessor
        sbp.StartData mTB, mContractSpec, mFrom, mTo, mNumber, mBarLength, mBarUnits, mSessionOnly
    Case FromFile
        Dim fbp As New FileBasedProcessor
        fbp.StartData mTB, mTickfileName, mFrom, mTo, mNumber, mBarLength, mBarUnits, mSessionOnly
    End Select
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
    Set gProcessor = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub processTimeframeCommand( _
                ByVal params As String)
Dim clp As CommandLineParser

Const ProcName As String = "processTimeframeCommand"
On Error GoTo Err

Set clp = CreateCommandLineParser(params, " ")

mBarLength = 0
mBarUnits = TimePeriodNone

If clp.NumberOfArgs < 1 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid timeframe - the bar length must be supplied"
    Exit Sub
End If

If Not IsInteger(clp.Arg(0), 1) Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid bar length '" & Trim$(clp.Arg(0)) & ": must be an integer > 0"
    Exit Sub
End If
mBarLength = CLng(clp.Arg(0))

mBarUnits = TimePeriodMinute
If Trim$(clp.Arg(1)) <> "" Then
    mBarUnits = TimePeriodUnitsFromString(clp.Arg(1))
    If mBarUnits = TimePeriodNone Then
        gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid bar units '" & Trim$(clp.Arg(1)) & ": must be one of s,m,h,d,w,mm,v,tv,tm"
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

If IsDate(params) Then
    mTo = CDate(params)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid to date '" & params & "'"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupCommonStudiesLib() As Boolean
Dim sl As Object
Const ProcName As String = "setupCommonStudiesLib"
On Error GoTo Err

Set sl = mTB.StudyLibraryManager.AddStudyLibrary(BuiltInStudyLibraryProgId, True, BuiltInStudyLibraryName)
If sl Is Nothing Then
    gCon.WriteErrorLine "Common studies library is not installed"
Else
    setupCommonStudiesLib = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupDbServiceProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupDbServiceProviders"
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

setupDbServiceProviders = True
    
On Error Resume Next
Dim sp As Object
Set sp = mTB.ServiceProviders.Add( _
                ProgId:="TBInfoBase27.ContractInfoSrvcProvider", _
                Enabled:=True, _
                ParamString:="Database Name=" & database & _
                            ";Database Type=" & dbtypeStr & _
                            ";Server=" & server & _
                            ";User name=" & username & _
                            ";Password=" & password, _
                Description:="Enable contract data from TradeBuild's database")
If sp Is Nothing Then
    gCon.WriteErrorLine "Required contract info service provider is not installed"
    setupDbServiceProviders = False
End If

Set sp = mTB.ServiceProviders.Add( _
                ProgId:="TBInfoBase27.HistDataServiceProvider", _
                Enabled:=True, _
                ParamString:="Database Name=" & database & _
                            ";Database Type=" & dbtypeStr & _
                            ";Server=" & server & _
                            ";User name=" & username & _
                            ";Password=" & password, _
                Description:="Enable historical bar data storage/retrieval to/from TradeBuild's database")
If sp Is Nothing Then
    gCon.WriteErrorLine "Required historical data service provider is not installed"
    setupDbServiceProviders = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupFileServiceProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupFileServiceProviders"
On Error GoTo Err

setupFileServiceProviders = True

mTickfileName = switchValue
    
On Error Resume Next
Dim sp As Object
Set sp = mTB.ServiceProviders.Add( _
                ProgId:="TickfileSP27.TickfileServiceProvider", _
                Enabled:=True, _
                ParamString:="Role=Input", _
                Description:="Historical tick data input from files")
If sp Is Nothing Then
    gCon.WriteErrorLine "Required tickfile service provider is not installed"
    setupFileServiceProviders = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function setupTwsServiceProviders( _
                ByVal switchValue As String) As Boolean
Const ProcName As String = "setupTwsServiceProviders"
On Error GoTo Err

setupTwsServiceProviders = True

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
    setupTwsServiceProviders = False
End If
    
If clientId = "" Then
    clientId = &H7A92DC3F
ElseIf Not IsInteger(clientId, 0) Then
    gCon.WriteErrorLine "Error: clientId must be an integer >= 0"
    setupTwsServiceProviders = False
End If
    
If setupTwsServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:=SPProgIdTwsRealtimeData, _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True;Role=Primary", _
                        Description:="Enable contract info from TWS"
End If

If setupTwsServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:=SPProgIdTwsContractData, _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True;Role=Primary", _
                        Description:="Enable contract info from TWS"
End If

If setupTwsServiceProviders Then
    mTB.ServiceProviders.Add _
                        ProgId:=SPProgIdTwsBarData, _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Enable historical data from TWS"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim failpoint As Long
Const ProcName As String = "setupServiceProviders"
On Error GoTo Err

setupServiceProviders = True

Set mTB = CreateTradeBuildAPI(SPRoleContractDataPrimary Or SPRoleHistoricalDataInput Or SPRoleTickfileInput Or SPRoleRealtimeData)

Select Case mSwitch
Case FromDb
    If Not setupDbServiceProviders(switchValue) Then setupServiceProviders = False
Case FromFile
    If Not setupFileServiceProviders(switchValue) Then setupServiceProviders = False
Case FromTws
    If Not setupTwsServiceProviders(switchValue) Then setupServiceProviders = False
End Select

If Not setupCommonStudiesLib Then setupServiceProviders = False

mTB.StartServiceProviders

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub showContractHelp()
gCon.WriteLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
End Sub

Private Sub showStdInHelp()
gCon.WriteLineToConsole "StdIn Format:"
gCon.WriteLineToConsole ""
gCon.WriteLineToConsole "#comment"
showContractHelp
gCon.WriteLineToConsole "from starttime"
gCon.WriteLineToConsole "to endtime"
gCon.WriteLineToConsole "number n               # -1 => return all available bars"
showTimeframeHelp
gCon.WriteLineToConsole "nonsess                # include bars outside session"
gCon.WriteLineToConsole "sess                   # include only bars within the session"
gCon.WriteLineToConsole "start"
gCon.WriteLineToConsole "stop"
End Sub

Private Sub showTimeframeHelp()
gCon.WriteLineToConsole "timeframe timeframespec"
gCon.WriteLineToConsole "  where"
gCon.WriteLineToConsole "    timeframespec  ::= length [units]"
gCon.WriteLineToConsole "    units          ::=     m   minutes (default)"
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






