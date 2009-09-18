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

Private Const ProjectName                   As String = "gcd"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const EchoCommand                   As String = "$ECHO"

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console
Private mCp As New ContractProcessor

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

Set clp = CreateCommandLineParser(Command)

If clp.Switch("?") Or _
    clp.NumberOfArgs > 0 Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf clp.Switch("fromdb") Then
    If setupDbServiceProvider(clp.switchValue("fromdb")) Then
        process
    End If
ElseIf clp.Switch("fromtws") Then
    If setupTwsServiceProvider(clp.switchValue("fromtws")) Then
        process
    End If
Else
    showUsage
End If

TradeBuildAPI.ServiceProviders.RemoveAll
TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then gCon.writeErrorLine Err.Description & " (" & Err.Source & ")"
TradeBuildAPI.ServiceProviders.RemoveAll
TerminateTWUtilities

    
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub process()
Dim inString As String
Dim lineNumber As Long

inString = Trim$(gCon.readLine(":"))
Do While inString <> gCon.eofString
    lineNumber = lineNumber + 1
    If inString = "" Then
        ' ignore blank lines
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments
    ElseIf Left$(inString, 1) = "$" Then
        ' process command
        If Len(inString) >= Len(EchoCommand) And _
            UCase$(Left$(inString, Len(EchoCommand))) = EchoCommand _
        Then
            gCon.writeLine Trim$(Right$(inString, Len(inString) - Len(EchoCommand)))
        Else
            gCon.writeErrorLine "Invalid command '" & Split(inString, " ")(0)
        End If
    Else
        processInput inString, lineNumber
    End If
    inString = Trim$(gCon.readLine(":"))
Loop
End Sub

Private Sub processInput( _
                ByVal inString As String, _
                ByVal lineNumber As Long)
' StdIn format:
' sectype,exchange,shortname,symbol,currency,expiry,strike,right,nametemplate

Dim validInput As Boolean
Dim sectype As SecurityTypes
Dim sectypeStr As String
Dim exchange As String
Dim shortname As String
Dim symbol As String
Dim currencyCode As String
Dim expiry As String
Dim strike As Double
Dim strikeStr As String
Dim optRight As OptionRights
Dim optRightStr As String
Dim nametemplate As String

Dim parser As CommandLineParser

Dim failpoint As Long
On Error GoTo Err

validInput = True

Set parser = CreateCommandLineParser(inString, ",")

sectypeStr = parser.Arg(0)
exchange = parser.Arg(1)
shortname = parser.Arg(2)
symbol = parser.Arg(3)
currencyCode = parser.Arg(4)
expiry = parser.Arg(5)
strikeStr = parser.Arg(6)
optRightStr = parser.Arg(7)
nametemplate = parser.Arg(8)

sectype = SecTypeFromString(sectypeStr)
If sectypeStr <> "" And sectype = SecTypeNone Then
    gCon.writeErrorLine "Line " & lineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validInput = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiry = Format(CDate(expiry), "yyyymmdd")
    ElseIf Len(expiry) = 6 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Right$(expiry, 2) & "/01") Then
            gCon.writeErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
            validInput = False
        End If
    ElseIf Len(expiry) = 8 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            gCon.writeErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
            validInput = False
        End If
    Else
        gCon.writeErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
        validInput = False
    End If
End If
            
If strikeStr <> "" Then
    If IsNumeric(strikeStr) Then
        strike = CDbl(strikeStr)
    Else
        gCon.writeErrorLine "Line " & lineNumber & ": Invalid strike '" & strikeStr & "'"
        validInput = False
    End If
End If

optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.writeErrorLine "Line " & lineNumber & ": Invalid right '" & optRightStr & "'"
    validInput = False
End If

        
If validInput Then
    mCp.process CreateContractSpecifier(shortname, _
                                        symbol, _
                                        exchange, _
                                        sectype, _
                                        currencyCode, _
                                        expiry, _
                                        strike, _
                                        optRight), _
                lineNumber, _
                nametemplate
End If

Exit Sub

Err:
gCon.writeErrorLine Err.Description
End Sub

Private Function setupDbServiceProvider( _
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

setupDbServiceProvider = True

On Error Resume Next
server = clp.Arg(0)
dbtypeStr = clp.Arg(1)
database = clp.Arg(2)
username = clp.Arg(3)
password = clp.Arg(4)
On Error GoTo 0

dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gCon.writeErrorLine "Error: invalid dbtype"
    setupDbServiceProvider = False
End If

If username <> "" And password = "" Then
    password = gCon.readLineFromConsole("Password:", "*")
End If
    
If setupDbServiceProvider Then
    TradeBuildAPI.ServiceProviders.Add _
                        ProgId:="TBInfoBase26.ContractInfoSrvcProvider", _
                        Enabled:=True, _
                        ParamString:="Database Name=" & database & _
                                    ";Database Type=" & dbtypeStr & _
                                    ";Server=" & server & _
                                    ";user name=" & username & _
                                    ";password=" & password, _
                        Description:="Enable contract data from TradeBuild's database"
    
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupDbServiceProvider = False

End Function

Private Function setupTwsServiceProvider( _
                ByVal switchValue As String) As Boolean
Dim tokens() As String
Dim server As String
Dim port As String
Dim clientId As String

Dim failpoint As Long
On Error GoTo Err

setupTwsServiceProvider = True

tokens = Split(switchValue, ",")

port = 7496
clientId = -1

On Error Resume Next
server = tokens(0)
port = tokens(1)
clientId = tokens(2)
On Error GoTo 0

If port <> "" Then
    If Not IsNumeric(port) Then
        gCon.writeErrorLine "Error: port must be numeric"
        setupTwsServiceProvider = False
    ElseIf port <= 0 Then
        gCon.writeErrorLine "Error: port must be > 0"
        setupTwsServiceProvider = False
    End If
End If
    
If clientId <> "" Then
    If Not IsNumeric(clientId) Then
        gCon.writeErrorLine "Error: clientId must be numeric"
        setupTwsServiceProvider = False
    End If
End If
    
If setupTwsServiceProvider Then
    TradeBuildAPI.ServiceProviders.Add _
                        ProgId:="IBTWSSP26.ContractInfoServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Enable contract data from TWS"
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupTwsServiceProvider = False

End Function

Private Sub showUsage()
gCon.writeErrorLine "Usage:"
gCon.writeErrorLine "gcd -fromdb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.writeErrorLine "    OR"
gCon.writeErrorLine "    -fromtws:[<twsserver>}[,[<port>][,[<clientid>]]]"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdIn Format:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "$echo text"
gCon.writeErrorLine "sectype,exchange,shortname,symbol,currency,expiry,strike,right"
End Sub
