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

Public Const ProjectName                    As String = "gxd27"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

Public gDb As TradingDB

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
    clp.NumberOfSwitches = 0 _
Then
    showUsage
Else
    If clp.Switch("fromdb") Then
        If setupDb(clp.switchValue("fromdb")) Then
            process
        End If
    Else
        showUsage
    End If
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then gCon.WriteErrorLine Err.Description
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
    Else
        processInput inString, lineNumber
    End If
    inString = Trim$(gCon.ReadLine(":"))
Loop
End Sub

Private Sub processexchange( _
                ByVal ex As Exchange)
gCon.WriteLine ex.Name & "," & ex.TimeZoneName & ",""" & ex.Notes & """"
End Sub

Private Sub processInput( _
                ByVal inString As String, _
                ByVal lineNumber As Long)
' StdIn format:
' [exchange]

Dim exF As ExchangeFactory
Dim summs As DataObjectSummaries
Dim summ As DataObjectSummary
Dim ex As Exchange
Dim fieldnames() As String

Dim failpoint As Long
On Error GoTo Err

Set exF = gDb.ExchangeFactory

If inString = "*" Then
    Set summs = exF.Query("", fieldnames)
    For Each summ In summs
        Set ex = exF.LoadByID(summ.Id)
        processexchange ex
    Next
Else
    Set ex = exF.LoadByName(inString)
    If ex Is Nothing Then
        gCon.WriteErrorLine "Line " & lineNumber & ": invalid exchange name '" & inString & "'"
    Else
        processexchange ex
    End If
End If

Exit Sub

Err:
gCon.WriteErrorLine Err.Description
End Sub

Private Function setupDb( _
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

setupDb = True

On Error Resume Next
server = clp.Arg(0)
dbtypeStr = clp.Arg(1)
database = clp.Arg(2)
username = clp.Arg(3)
password = clp.Arg(4)
On Error GoTo 0

If username <> "" And password = "" Then
    password = gCon.ReadLineFromConsole("Password:", "*")
End If
    
dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gCon.WriteErrorLine "Error: invalid dbtype"
    setupDb = False
End If
    
If setupDb Then
    Set gDb = CreateTradingDB(GenerateConnectionString(dbtype, _
                                                        server, _
                                                        database, _
                                                        username, _
                                                        password), _
                            dbtype)
End If

Exit Function

Err:
gCon.WriteErrorLine Err.Description
setupDb = False

End Function

Private Sub showUsage()
gCon.WriteErrorLine "Usage:"
gCon.WriteErrorLine "gxd27 -fromdb:databaseserver,databasetype,catalog[,username[,password]]"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "StdIn Format:"
gCon.WriteErrorLine "*"
gCon.WriteErrorLine "exchange"
End Sub




