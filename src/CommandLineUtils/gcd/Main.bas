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
Private Const ModuleName                    As String = "Main"

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

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
    
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub getContracts( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal lineNumber As Long)
Dim cp As New ContractProcessor
cp.process contractSpec, lineNumber
End Sub

Private Sub process()
Dim inString As String
Dim lineNumber As Long

inString = gCon.readLine
Do While inString <> gCon.eofString
    lineNumber = lineNumber + 1
    processInput inString, lineNumber
    inString = gCon.readLine
Loop
End Sub

Private Sub processInput( _
                ByVal inString As String, _
                ByVal lineNumber As Long)
' StdIn format:
' sectype,exchange,shortname,symbol,currency,expiry,strike,right

Dim validInput As Boolean
Dim tokens() As String
Dim sectype As SecurityTypes
Dim sectypeStr As String
Dim exchange As String
Dim shortname As String
Dim symbol As String
Dim currencyCode As String
Dim expiry As Date
Dim strike As Double
Dim strikeStr As String
Dim optRight As OptionRights
Dim optRightStr As String

Dim failpoint As Long
On Error GoTo Err

validInput = True

tokens = Split(inString, ",")

On Error Resume Next
sectypeStr = tokens(0)
exchange = tokens(1)
shortname = tokens(2)
symbol = tokens(3)
currencyCode = tokens(4)
expiry = tokens(5)
strikeStr = tokens(6)
optRightStr = tokens(7)
On Error GoTo 0

sectype = SecTypeFromString(sectypeStr)
If sectypeStr <> "" And sectype = SecTypeNone Then
    gCon.writeErrorLine "Line " & lineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validInput = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiry = Format(CDate(expiry), "yyyymmdd")
    ElseIf Len(expiry) = 6 Then
        If Not IsDate(Left$(expiry, 4) & "/" & right$(expiry, 2) & "/01") Then
            gCon.writeErrorLine "Line " & lineNumber & ": Invalid expiry '" & expiry & "'"
            validInput = False
        End If
    ElseIf Len(expiry) = 8 Then
        If Not IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 4, 2) & "/" & right$(expiry, 2)) Then
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
    getContracts CreateContractSpecifier(shortname, symbol, exchange, SecTypeCash, currencyCode, expiry, strike, optRight), lineNumber
End If

Exit Sub

Err:
gCon.writeErrorLine Err.Description
End Sub

Private Function setupDbServiceProvider( _
                ByVal switchValue As String) As Boolean

End Function

Private Function setupTwsServiceProvider( _
                ByVal switchValue As String) As Boolean
Dim tokens() As String
Dim server As String
Dim port As String
Dim clientId As String

Dim failpoint As Long
On Error GoTo Err

tokens = Split(switchValue, ",")

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
    

TradeBuildAPI.ServiceProviders.Add _
                    ProgId:="IBTWSSP26.ContractInfoServiceProvider", _
                    Enabled:=True, _
                    ParamString:="Server=" & tokens(0) & _
                                ";Port=" & tokens(1) & _
                                ";Client Id=" & clientId & _
                                ";Provider Key=IB;Keep Connection=True", _
                    logLevel:=LogLevelLow, _
                    Description:="Enable contract data from TWS"

setupTwsServiceProvider = True

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupTwsServiceProvider = False

End Function

Private Sub showUsage()
gCon.writeLine "Usage:"
gCon.writeLine "gcd -fromdb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.writeLine "    -fromtws:[<twsserver>[,[<port>]][,[<clientid>]]"
End Sub
