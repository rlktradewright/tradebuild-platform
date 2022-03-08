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

Public Const ProjectName                    As String = "gtd"
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

Private mDBClient As DBClient

Private mLineNumber As Long
Private mContractSpec As IContractSpecifier
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

On Error GoTo Err

InitialiseTWUtilities

ApplicationGroupName = "TradeWright"
ApplicationName = "GTD v" & App.Major & "." & App.Minor
SetupDefaultLogging command

Set gCon = GetConsole


Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(command)

If clp.Switch("?") Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf Not clp.Switch("fromdb") Then
    showUsage
ElseIf Not setupServiceProviders(clp.switchValue("fromdb")) Then

ElseIf Not clp.Switch("Speed") Then
    process
ElseIf Not IsInteger(clp.switchValue("Speed")) Then
    gCon.WriteErrorLine "Speed must be an integer"
Else
    mSpeed = CLng(clp.switchValue("Speed"))
    process
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then
    gCon.WriteErrorLine Err.Description
    If Err.Source <> "" Then gCon.WriteErrorLine Err.Source
End If
TerminateTWUtilities
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub process()
Dim inString As String
Dim command As String
Dim params As String

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
        Case PauseCommand
            processPauseCommand
        Case StopCommand
            processStopCommand
        Case SpeedCommand
            processSpeedCommand params
        Case RawCommand
            processRawCommand params
        Case Else
            gCon.WriteErrorLine "Invalid command '" & command & "'"
        End Select
    End If
    inString = Trim$(gCon.ReadLine(":"))
Loop
End Sub

Private Sub processContractCommand( _
                ByVal params As String)
'params: shortname,sectype,exchange,symbol,currency,expiry.multiplier,strike,right

If gProcessor Is Nothing Then
ElseIf gProcessor.IsRunning Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot set contract - already running"
    Exit Sub
End If

Dim clp As CommandLineParser: Set clp = CreateCommandLineParser(params, InputSep)

Dim validParams As Boolean: validParams = True

Dim sectypeStr As String: sectypeStr = Trim$(clp.Arg(1))
Dim exchange As String: exchange = Trim$(clp.Arg(2))
Dim shortname As String: shortname = Trim$(clp.Arg(0))
Dim symbol As String: symbol = Trim$(clp.Arg(3))
Dim currencyCode As String: currencyCode = Trim$(clp.Arg(4))
Dim expiry As String: expiry = Trim$(clp.Arg(5))
Dim multiplierStr: multiplierStr = Trim$(clp.Arg(6))
Dim strikeStr As String: strikeStr = Trim$(clp.Arg(7))
Dim optRightStr As String: optRightStr = Trim$(clp.Arg(8))

Dim sectype As SecurityTypes: sectype = SecTypeFromString(sectypeStr)
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
    multiplier = 0#
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
    gProcessor.SetContract mContractSpec
End If
End Sub

Private Sub processFromCommand( _
                ByVal params As String)
If IsDate(params) Then
    mFrom = CDate(params)
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid from date '" & params & "'"
End If
End Sub

Private Sub processPauseCommand()
If gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot pause - not started"
ElseIf gProcessor.IsPaused Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Already paused"
Else
    gProcessor.PauseData
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
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid raw command '" & params & "'"
End Select
End Sub

Private Sub processSpeedCommand( _
                ByVal params As String)
If IsInteger(params) Then
    mSpeed = CLng(params)
    If Not gProcessor Is Nothing Then
        gProcessor.Speed = mSpeed
    End If
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid Speed factor '" & params & "'"
End If
End Sub


Private Sub processStartCommand()
If mContractSpec Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - no contract specified"
ElseIf mFrom = 0 Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - from time not specified"
ElseIf gProcessor.IsPaused Then
    gProcessor.ResumeData
ElseIf Not gProcessor.IsRunning Then
    gProcessor.StartData mFrom, mTo, mSpeed, mRaw
Else
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot start - already running"
End If
End Sub

Private Sub processStopCommand()
If gProcessor Is Nothing Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": Cannot stop - not started"
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
    gCon.WriteErrorLine "Line " & mLineNumber & ": Invalid to date '" & params & "'"
End If
End Sub

Private Function setupServiceProviders( _
                ByVal switchValue As String) As Boolean
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupServiceProviders = True

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
    setupServiceProviders = False
End If
    
If setupServiceProviders Then
    Set mDBClient = CreateTradingDBClient(dbtype, server, database, username, password, True)
    If mDBClient Is Nothing Then
        gCon.WriteErrorLine "Error: can't access database"
        setupServiceProviders = False
    Else
        Set gProcessor = New Processor
        gProcessor.Initialise mDBClient
    End If
End If

Exit Function

Err:
gCon.WriteErrorLine Err.Description
setupServiceProviders = False
End Function

Private Sub showUsage()
gCon.WriteErrorLine "Usage:"
gCon.WriteErrorLine "gtd27 -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.WriteErrorLine "    -Speed:n"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "StdIn Format:"
gCon.WriteErrorLine "#comment"
gCon.WriteErrorLine "contract shortname,sectype,exchange,symbol,currency,expiry,multiplier,strike,right"
gCon.WriteErrorLine "from starttime"
gCon.WriteErrorLine "to endtime"
gCon.WriteErrorLine "speed n"
gCon.WriteErrorLine "raw [on | off]"
gCon.WriteErrorLine "start"
gCon.WriteErrorLine "pause"
gCon.WriteErrorLine "stop"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "StdOUt Format:"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "timestamp, ticktype, tickvalues"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "  where"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "    timestamp ::= yyyy-mm-dd hh:mm:ss.nnn"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "    ticktype ::=   B   bid"
gCon.WriteErrorLine "                   A   ask"
gCon.WriteErrorLine "                   T   trade"
gCon.WriteErrorLine "                   V   volume"
gCon.WriteErrorLine "                   I   open interest"
gCon.WriteErrorLine "                   O   open"
gCon.WriteErrorLine "                   H   high"
gCon.WriteErrorLine "                   L   low"
gCon.WriteErrorLine "                   C   previous session close"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "    tickvalues ::=  price[,size][,+ | - | =][,+ | - | =]"
gCon.WriteErrorLine ""
End Sub




