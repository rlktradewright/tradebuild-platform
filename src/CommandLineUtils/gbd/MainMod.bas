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

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

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

Set gCon = GetConsole

Set clp = CreateCommandLineParser(command)

If clp.Switch("?") Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf Not clp.Switch("fromdb") Then
    showUsage
ElseIf Not setupServiceProviders(clp.switchValue("fromdb")) Then
Else
    process
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
        Case Else
            gCon.writeErrorLine "Invalid command '" & command
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

If Not gProcessor Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot set contract - already running"
    Exit Sub
End If

Set clp = CreateCommandLineParser(params, InputSep)

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
If Not IsInteger(params, 1) Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid number '" & params & "'" & ": must be an integer > 0"
Else
    mNumber = CLng(params)
End If
End Sub

Private Sub processSessCommand()
mSessionOnly = True
End Sub

Private Sub processStartCommand()
If mContractSpec Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mFrom = 0 And mNumber = 0 Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - either from time or number of bars must be specified"
ElseIf mBarUnits = TimePeriodNone Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - timeframe not specified"
ElseIf gProcessor Is Nothing Then
    Set gProcessor = New Processor
    gProcessor.startData mContractSpec, mFrom, mTo, mNumber, mBarLength, mBarUnits, mSessionOnly
Else
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - already running"
End If
End Sub

Private Sub processStopCommand()
If gProcessor Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot stop - not started"
Else
    gProcessor.stopData
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

Private Function setupServiceProviders( _
                ByVal switchValue As String) As Boolean
Dim clp As CommandLineParser
Dim server As String
Dim dbtypeStr As String
Dim dbtype As DatabaseTypes
Dim database As String
Dim username As String
Dim password As String
Dim sp As Object

Dim failpoint As Long
On Error GoTo Err

Set clp = CreateCommandLineParser(switchValue, ",")

setupServiceProviders = True

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
    setupServiceProviders = False
End If
    
If setupServiceProviders Then
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
        setupServiceProviders = False
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
        setupServiceProviders = False
    End If
End If

Set sp = AddStudyLibrary("CmnStudiesLib26.StudyLib", True, "Built-in")
If sp Is Nothing Then
    gCon.writeErrorLine "Common studies library is not installed"
    setupServiceProviders = False
End If


Exit Function

Err:
gCon.writeErrorLine Err.Description
setupServiceProviders = False

End Function

Private Sub showUsage()
gCon.writeErrorLine "Usage:"
gCon.writeErrorLine "gbd -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdIn Format:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
gCon.writeErrorLine "from starttime"
gCon.writeErrorLine "to endtime"
gCon.writeErrorLine "number n"
gCon.writeErrorLine "timeframe timeframespec"
gCon.writeErrorLine "nonsess"
gCon.writeErrorLine "sess"
gCon.writeErrorLine "start"
gCon.writeErrorLine "stop"
gCon.writeErrorLine ""
gCon.writeErrorLine "  where"
gCon.writeErrorLine ""
gCon.writeErrorLine "    timeframespec  ::= length [units]"
gCon.writeErrorLine ""
gCon.writeErrorLine "    units          ::=     m   minutes (default)"
gCon.writeErrorLine "                           h   hours"
gCon.writeErrorLine "                           d   days"
gCon.writeErrorLine "                           w   weeks"
gCon.writeErrorLine "                           mm   months"
gCon.writeErrorLine "                           v   volume"
gCon.writeErrorLine "                           tv  tick volume"
gCon.writeErrorLine "                           tm   ticks movement"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdOUt Format:"
gCon.writeErrorLine ""
gCon.writeErrorLine "timestamp , ticktype, tickvalues"
gCon.writeErrorLine ""
gCon.writeErrorLine "  where"
gCon.writeErrorLine ""
gCon.writeErrorLine "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.writeErrorLine ""
gCon.writeErrorLine ""
End Sub




