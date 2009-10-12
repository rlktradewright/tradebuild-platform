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

Private Const ProjectName                   As String = "gbd"
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

Public gCon As Console

Private mSwitch As Switches

Private mTickfileName As String

Private mLineNumber As Long
Private mContractSpec As ContractSpecifier
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

Public Sub Main()
Dim clp As CommandLineParser

On Error GoTo Err

InitialiseTWUtilities
'EnableTracing "tradebuild"
'EnableTracing "tickfilesp"

mTo = MaxDate
mNumber = -1

Set gCon = GetConsole

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
If Not gCon Is Nothing Then gCon.writeErrorLine Err.Description
TerminateTWUtilities

    
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub process()
Dim inString As String
Dim command As String
Dim params As String

inString = Trim$(gCon.readLine(":"))
Do While inString <> gCon.eofString
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
            gCon.writeErrorLine "Invalid command '" & command & "'"
        End Select
    End If
    inString = Trim$(gCon.readLine(":"))
Loop
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
    gCon.writeLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
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
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validParams = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiry = Format(CDate(expiry), "yyyymmdd")
    ElseIf Len(expiry) = 6 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Right$(expiry, 2) & "/01") Then
            gCon.writeErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
            validParams = False
        End If
    ElseIf Len(expiry) = 8 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            gCon.writeErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
            validParams = False
        End If
    Else
        gCon.writeErrorLine "Line " & mLineNumber & ": Invalid expiry '" & expiry & "'"
        validParams = False
    End If
End If
            
If strikeStr <> "" Then
    If IsNumeric(strikeStr) Then
        strike = CDbl(strikeStr)
    Else
        gCon.writeErrorLine "Line " & mLineNumber & ": Invalid strike '" & strikeStr & "'"
        validParams = False
    End If
End If

optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid right '" & optRightStr & "'"
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

Exit Sub

Err:
Set mContractSpec = Nothing
gCon.writeErrorLine "Error: " & Err.Description
End Sub

Private Sub processFromCommand( _
                ByVal params As String)
If IsDate(params) Then
    mFrom = CDate(params)
Else
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid from date '" & params & "'"
End If
End Sub

Private Sub processNonSessCommand()
mSessionOnly = False
End Sub

Private Sub processNumberCommand( _
                ByVal params As String)
If Not IsInteger(params, 1) And params <> "-1" Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid number '" & params & "'" & ": must be an integer > 0"
Else
    mNumber = CLng(params)
    If mSwitch = FromFile Then gCon.writeLineToConsole "number command is ignored for tickfile input"
End If
End Sub

Private Sub processSessCommand()
mSessionOnly = True
End Sub

Private Sub processStartCommand()
If mSwitch <> FromFile And mContractSpec Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mSwitch <> FromFile And mFrom = 0 And mNumber = 0 Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - either from time or number of bars must be specified"
ElseIf mBarUnits = TimePeriodNone Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - timeframe not specified"
ElseIf Not gProcessor Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - already running"
Else
       
    Set gProcessor = New Processor
    
    Select Case mSwitch
    Case FromDb, FromTws
        Dim sbp As New StreamBasedProcessor
        sbp.startData mContractSpec, mFrom, mTo, mNumber, mBarLength, mBarUnits, mSessionOnly
    Case FromFile
        Dim fbp As New FileBasedProcessor
        fbp.startData mTickfileName, mFrom, mTo, mNumber, mBarLength, mBarUnits, mSessionOnly
    End Select
End If
End Sub

Private Sub processStopCommand()
If gProcessor Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot stop - not started"
Else
    gProcessor.StopData
    Set gProcessor = Nothing
End If
End Sub

Private Sub processTimeframeCommand( _
                ByVal params As String)
Dim clp As CommandLineParser

Set clp = CreateCommandLineParser(params, " ")

mBarLength = 0
mBarUnits = TimePeriodNone

If clp.NumberOfArgs < 1 Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid timeframe - the bar length must be supplied"
    Exit Sub
End If

If Not IsInteger(clp.Arg(0), 1) Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid bar length '" & Trim$(clp.Arg(0)) & ": must be an integer > 0"
    Exit Sub
End If
mBarLength = CLng(clp.Arg(0))

mBarUnits = TimePeriodMinute
If Trim$(clp.Arg(1)) <> "" Then
    mBarUnits = TimePeriodUnitsFromString(clp.Arg(1))
    If mBarUnits = TimePeriodNone Then
        gCon.writeErrorLine "Line " & mLineNumber & ": Invalid bar units '" & Trim$(clp.Arg(1)) & ": must be one of s,m,h,d,w,mm,v,tv,tm"
    Exit Sub
    End If
End If

End Sub

Private Sub processToCommand( _
                ByVal params As String)
If IsDate(params) Then
    mTo = CDate(params)
Else
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid to date '" & params & "'"
End If
End Sub

Private Function setupCommonStudiesLib() As Boolean
Dim sl As Object
Set sl = AddStudyLibrary("CmnStudiesLib26.StudyLib", True, "Built-in")
If sl Is Nothing Then
    gCon.writeErrorLine "Common studies library is not installed"
Else
    setupCommonStudiesLib = True
End If
End Function

Private Function setupDbServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim clp As CommandLineParser
Dim server As String
Dim dbtypeStr As String
Dim dbtype As DatabaseTypes
Dim database As String
Dim username As String
Dim password As String
Dim sp As Object

Set clp = CreateCommandLineParser(switchValue, ",")

On Error Resume Next
server = clp.Arg(0)
dbtypeStr = clp.Arg(1)
database = clp.Arg(2)
username = clp.Arg(3)
password = clp.Arg(4)
On Error GoTo 0

If username <> "" And password = "" Then
    password = gCon.readLineFromConsole("Password:", "*")
End If
    
dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gCon.writeErrorLine "Error: invalid dbtype"
    Exit Function
End If

setupDbServiceProviders = True
    
On Error Resume Next
Set sp = TradeBuildAPI.ServiceProviders.Add( _
                ProgId:="TBInfoBase26.ContractInfoSrvcProvider", _
                Enabled:=True, _
                ParamString:="Database Name=" & database & _
                            ";Database Type=" & dbtypeStr & _
                            ";Server=" & server & _
                            ";User name=" & username & _
                            ";Password=" & password, _
                Description:="Enable contract data from TradeBuild's database")
If sp Is Nothing Then
    gCon.writeErrorLine "Required contract info service provider is not installed"
    setupDbServiceProviders = False
End If

Set sp = TradeBuildAPI.ServiceProviders.Add( _
                ProgId:="TBInfoBase26.HistDataServiceProvider", _
                Enabled:=True, _
                ParamString:="Database Name=" & database & _
                            ";Database Type=" & dbtypeStr & _
                            ";Server=" & server & _
                            ";User name=" & username & _
                            ";Password=" & password, _
                Description:="Enable historical bar data storage/retrieval to/from TradeBuild's database")
If sp Is Nothing Then
    gCon.writeErrorLine "Required historical data service provider is not installed"
    setupDbServiceProviders = False
End If

End Function

Private Function setupFileServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim sp As Object

setupFileServiceProviders = True

mTickfileName = switchValue
    
On Error Resume Next
Set sp = TradeBuildAPI.ServiceProviders.Add( _
                ProgId:="TickfileSP26.TickfileServiceProvider", _
                Enabled:=True, _
                ParamString:="Access mode=ReadOnly", _
                Description:="Historical tick data input from files")
If sp Is Nothing Then
    gCon.writeErrorLine "Required tickfile service provider is not installed"
    setupFileServiceProviders = False
End If

End Function

Private Function setupTwsServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim clp As CommandLineParser
Dim server As String
Dim port As String
Dim clientId As String

On Error GoTo Err

Set clp = CreateCommandLineParser(switchValue, ",")

setupTwsServiceProviders = True

port = 7496
clientId = &H7A92DC3F

On Error Resume Next
server = clp.Arg(0)
port = clp.Arg(1)
clientId = clp.Arg(2)
On Error GoTo 0

If port <> "" Then
    If Not IsInteger(port, 0) Then
        gCon.writeErrorLine "Error: port must be a positive integer"
        setupTwsServiceProviders = False
    End If
End If
    
If clientId <> "" Then
    If Not IsInteger(clientId) Then
        gCon.writeErrorLine "Error: clientId must be an integer"
        setupTwsServiceProviders = False
    End If
End If
    
If setupTwsServiceProviders Then
    TradeBuildAPI.ServiceProviders.Add _
                        ProgId:="IBTWSSP26.ContractInfoServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Enable contract info from TWS"
End If

If setupTwsServiceProviders Then
    TradeBuildAPI.ServiceProviders.Add _
                        ProgId:="IBTWSSP26.HistDataServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Enable historical data from TWS"
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupTwsServiceProviders = False

End Function

Private Function setupServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim failpoint As Long
On Error GoTo Err

setupServiceProviders = True

Select Case mSwitch
Case FromDb
    If Not setupDbServiceProviders(switchValue) Then setupServiceProviders = False
Case FromFile
    If Not setupFileServiceProviders(switchValue) Then setupServiceProviders = False
Case FromTws
    If Not setupTwsServiceProviders(switchValue) Then setupServiceProviders = False
End Select

If Not setupCommonStudiesLib Then setupServiceProviders = False

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupServiceProviders = False

End Function

Private Sub showContractHelp()
gCon.writeLineToConsole "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
End Sub

Private Sub showStdInHelp()
gCon.writeLineToConsole "StdIn Format:"
gCon.writeLineToConsole ""
gCon.writeLineToConsole "#comment"
showContractHelp
gCon.writeLineToConsole "from starttime"
gCon.writeLineToConsole "to endtime"
gCon.writeLineToConsole "number n               # -1 => return all available bars"
showTimeframeHelp
gCon.writeLineToConsole "nonsess                # include bars outside session"
gCon.writeLineToConsole "sess                   # include only bars within the session"
gCon.writeLineToConsole "start"
gCon.writeLineToConsole "stop"
End Sub

Private Sub showTimeframeHelp()
gCon.writeLineToConsole "timeframe timeframespec"
gCon.writeLineToConsole "  where"
gCon.writeLineToConsole "    timeframespec  ::= length [units]"
gCon.writeLineToConsole "    units          ::=     m   minutes (default)"
gCon.writeLineToConsole "                           h   hours"
gCon.writeLineToConsole "                           d   days"
gCon.writeLineToConsole "                           w   weeks"
gCon.writeLineToConsole "                           mm   months"
gCon.writeLineToConsole "                           v   volume (constant volume bars)"
gCon.writeLineToConsole "                           tv  tick volume (constant tick volume bars)"
gCon.writeLineToConsole "                           tm   ticks movement (constant range bars)"
End Sub

Private Sub showUsage()
gCon.writeLineToConsole "Usage:"
gCon.writeLineToConsole "gbd -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.writeLineToConsole "    OR"
gCon.writeLineToConsole "    -fromfile:tickfilepath"
gCon.writeLineToConsole "    OR"
gCon.writeLineToConsole "    -fromtws: [twsserver] [,[port][,[clientid]]]"
gCon.writeLineToConsole ""
showStdInHelp
gCon.writeLineToConsole ""
gCon.writeLineToConsole "StdOut Format:"
gCon.writeLineToConsole ""
gCon.writeLineToConsole "timestamp,open,high,low,close,volume,tickvolume"
gCon.writeLineToConsole ""
gCon.writeLineToConsole "  where"
gCon.writeLineToConsole ""
gCon.writeLineToConsole "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.writeLineToConsole ""
gCon.writeLineToConsole ""
End Sub




