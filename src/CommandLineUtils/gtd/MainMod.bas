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

Private Const ProjectName                   As String = "gtd"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const ContractCommand               As String = "CONTRACT"
Private Const FromCommand                   As String = "FROM"
Private Const ToCommand                     As String = "TO"
Private Const StartCommand                  As String = "START"
Private Const PauseCommand                  As String = "PAUSE"
Private Const StopCommand                   As String = "STOP"
Private Const SpeedCommand                  As String = "SPEED"
Private Const RawCommand                    As String = "RAW"

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

Private mLineNumber As Long
Private mContractSpec As ContractSpecifier
Private mFrom As Date
Private mTo As Date

Private mSpeed As Long

Private mRaw As Boolean

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

ElseIf Not clp.Switch("speed") Then
    process
ElseIf Not IsInteger(clp.switchValue("speed")) Then
    gCon.writeErrorLine "Speed must be an integer"
Else
    mSpeed = CLng(clp.switchValue("speed"))
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
        Case PauseCommand
            processPauseCommand
        Case StopCommand
            processStopCommand
        Case SpeedCommand
            processSpeedCommand params
        Case RawCommand
            processRawCommand params
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

Private Sub processPauseCommand()
If gProcessor Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot pause - not started"
ElseIf gProcessor.tickerState <> TickerStatePaused Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Already paused"
Else
    gProcessor.stopData
    Set gProcessor = Nothing
End If
End Sub

Private Sub processRawCommand( _
                ByVal params As String)
Select Case UCase$(Trim$(params))
Case ""
    mRaw = Not mRaw
Case "ON", "YES", "TRUE", "T"
    mRaw = True
Case "OFF", "NO", "FALSE", "F"
    mRaw = True
Case Else
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid raw command '" & params & "'"
End Select
End Sub

Private Sub processSpeedCommand( _
                ByVal params As String)
If IsInteger(params) Then
    mSpeed = CLng(params)
    If Not gProcessor Is Nothing Then
        gProcessor.speed = mSpeed
    End If
Else
    gCon.writeErrorLine "Line " & mLineNumber & ": Invalid speed factor '" & params & "'"
End If
End Sub


Private Sub processStartCommand()
If mContractSpec Is Nothing Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mFrom = 0 Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Cannot start - from time not specified"
ElseIf gProcessor Is Nothing Then
    Set gProcessor = New Processor
    gProcessor.startData mContractSpec, mFrom, mTo, mSpeed, mRaw
ElseIf gProcessor.tickerState = TickerStatePaused Then
    gProcessor.resumeData
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
    Dim sp As Object
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
                    ProgId:="TBInfoBase26.TickfileServiceProvider", _
                    Enabled:=True, _
                    ParamString:="Database Name=" & database & _
                                ";Database Type=" & dbtypeStr & _
                                ";Server=" & server & _
                                ";User name=" & username & _
                                ";Password=" & password, _
                    Description:="Enable historical tick data storage/retrieval to/from TradeBuild's database")
    If sp Is Nothing Then
        gCon.writeErrorLine "Required tickfile service provider is not installed"
        setupServiceProviders = False
    End If
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupServiceProviders = False

End Function

Private Sub showUsage()
gCon.writeErrorLine "Usage:"
gCon.writeErrorLine "gtd -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.writeErrorLine "    -speed:n"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdIn Format:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "contract shortname,sectype,exchange,symbol,currency,expiry,strike,right"
gCon.writeErrorLine "from starttime"
gCon.writeErrorLine "to endtime"
gCon.writeErrorLine "speed n"
gCon.writeErrorLine "raw [on | off]"
gCon.writeErrorLine "start"
gCon.writeErrorLine "pause"
gCon.writeErrorLine "stop"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdOUt Format:"
gCon.writeErrorLine ""
gCon.writeErrorLine "timestamp, ticktype, tickvalues"
gCon.writeErrorLine ""
gCon.writeErrorLine "  where"
gCon.writeErrorLine ""
gCon.writeErrorLine "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.writeErrorLine ""
gCon.writeErrorLine "    ticktype ::=   B   bid"
gCon.writeErrorLine "                   A   ask"
gCon.writeErrorLine "                   T   trade"
gCon.writeErrorLine "                   V   volume"
gCon.writeErrorLine "                   I   open interest"
gCon.writeErrorLine "                   O   open"
gCon.writeErrorLine "                   H   high"
gCon.writeErrorLine "                   L   low"
gCon.writeErrorLine "                   C   previous session close"
gCon.writeErrorLine ""
gCon.writeErrorLine "    tickvalues ::=  price[,size][,+ | - | =][,+ | - | =]"
gCon.writeErrorLine ""
End Sub




