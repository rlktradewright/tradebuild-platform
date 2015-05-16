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

Public Const ProjectName                    As String = "gcd27"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const EchoCommand                   As String = "$ECHO"

'@================================================================================
' Member variables
'@================================================================================

Public gCon                                         As Console
Private mCp                                         As ContractProcessor

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

Public Sub Main()
On Error GoTo Err

InitialiseTWUtilities

Set gCon = GetConsole

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(Command)

If clp.Switch("?") Or _
    clp.NumberOfArgs > 0 Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
    Exit Sub
End If

Set mTB = CreateTradeBuildAPI(SPRoleContractDataPrimary)
Set mCp = New ContractProcessor
mCp.initialise mTB

If clp.Switch("fromdb") Then
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

mTB.ServiceProviders.RemoveAll
TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then gCon.WriteErrorLine Err.Description & " (" & Err.Source & ")"
mTB.ServiceProviders.RemoveAll
TerminateTWUtilities

    
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub process()
Dim inString As String
Dim lineNumber As Long

inString = Trim$(gCon.ReadLine(":"))
Do While inString <> gCon.EofString
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
            gCon.WriteLine Trim$(Right$(inString, Len(inString) - Len(EchoCommand)))
        Else
            gCon.WriteErrorLine "Invalid command '" & Split(inString, " ")(0)
        End If
    Else
        processInput inString, lineNumber
    End If
    inString = Trim$(gCon.ReadLine(":"))
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
    gCon.WriteErrorLine "Line " & lineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validInput = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiry = Format(CDate(expiry), "yyyymmdd")
    ElseIf Len(expiry) = 6 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Right$(expiry, 2) & "/01") Then
            gCon.WriteErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
            validInput = False
        End If
    ElseIf Len(expiry) = 8 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            gCon.WriteErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
            validInput = False
        End If
    Else
        gCon.WriteErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
        validInput = False
    End If
End If
            
If strikeStr <> "" Then
    If IsNumeric(strikeStr) Then
        strike = CDbl(strikeStr)
    Else
        gCon.WriteErrorLine "Line " & lineNumber & ": Invalid strike '" & strikeStr & "'"
        validInput = False
    End If
End If

optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.WriteErrorLine "Line " & lineNumber & ": Invalid right '" & optRightStr & "'"
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
gCon.WriteErrorLine Err.Description
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
    gCon.WriteErrorLine "Error: invalid dbtype"
    setupDbServiceProvider = False
End If

If username <> "" And password = "" Then
    password = gCon.ReadLineFromConsole("Password:", "*")
End If
    
If setupDbServiceProvider Then
    mTB.ServiceProviders.Add _
                        ProgId:="TBInfoBase27.ContractInfoSrvcProvider", _
                        Enabled:=True, _
                        ParamString:="Role=PRIMARY" & _
                                    ";Database Name=" & database & _
                                    ";Database Type=" & dbtypeStr & _
                                    ";Server=" & server & _
                                    ";user name=" & username & _
                                    ";password=" & password, _
                        Description:="Enable contract data from TradeBuild's database"
    mTB.StartServiceProviders
End If

Exit Function

Err:
gCon.WriteErrorLine Err.Description
setupDbServiceProvider = False

End Function

Private Function setupTwsServiceProvider( _
                ByVal switchValue As String) As Boolean
On Error GoTo Err

Dim clp As CommandLineParser
Set clp = CreateCommandLineParser(switchValue, ",")

setupTwsServiceProvider = True

On Error Resume Next
Dim server As String
server = clp.Arg(0)

Dim port As String
port = clp.Arg(1)

Dim clientId As String
clientId = clp.Arg(2)
On Error GoTo Err

If port = "" Then
    port = "7496"
ElseIf Not IsInteger(port, 1) Then
        gCon.WriteErrorLine "Error: port must be a positive integer > 0"
        setupTwsServiceProvider = False
End If
    
If clientId = "" Then
    clientId = "1952361208"
ElseIf Not IsInteger(clientId, 0) Then
        gCon.WriteErrorLine "Error: clientId must be an integer >= 0"
        setupTwsServiceProvider = False
End If
    
If setupTwsServiceProvider Then
    mTB.ServiceProviders.Add _
                        ProgId:="IBTWSSP27.ContractInfoServiceProvider", _
                        Enabled:=True, _
                        ParamString:="Role=PRIMARY" & _
                                    ";Server=" & server & _
                                    ";Port=" & port & _
                                    ";Client Id=" & clientId & _
                                    ";Provider Key=IB;Keep Connection=True", _
                        Description:="Enable contract data from TWS"
    mTB.StartServiceProviders
End If

Exit Function

Err:
gCon.WriteErrorLine Err.Description
setupTwsServiceProvider = False

End Function

Private Sub showUsage()
gCon.WriteErrorLine "Usage:"
gCon.WriteErrorLine "gcd -fromdb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.WriteErrorLine "    OR"
gCon.WriteErrorLine "    -fromtws:[<twsserver>}[,[<port>][,[<clientid>]]]"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "StdIn Format:"
gCon.WriteErrorLine "#comment"
gCon.WriteErrorLine "$echo text"
gCon.WriteErrorLine "sectype,exchange,shortname,symbol,currency,expiry,strike,right"
End Sub
