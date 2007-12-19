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

Private Const ProjectName                   As String = "uxd"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

Public gDb As TradingDB
Public gInstrumentClass As InstrumentClass

' if set, existing records are to be updated
Public gUpdate As Boolean

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
    If clp.Switch("todb") Then
        If setupDb(clp.switchValue("todb")) Then
            If clp.Switch("U") Then gUpdate = True
            process
        End If
    Else
        showUsage
    End If
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
Dim lineNumber As Long

inString = Trim$(gCon.readLine(":"))
Do While inString <> gCon.eofString
    lineNumber = lineNumber + 1
    If inString = "" Then
        ' ignore blank lines
    ElseIf Left$(inString, 1) = "#" Then
        ' ignore comments
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
' name,shortname,symbol,expiry,strike,right[,[sectype][,[exchange][,[currency][,[ticksize][,[tickvalue]]]]]]
Dim validInput As Boolean
Dim tokens() As String
Dim name As String
Dim timezone As String
Dim notes As String
Dim update As Boolean

Dim failpoint As Long
On Error GoTo Err

validInput = True

tokens = Split(inString, InputSep)

On Error Resume Next
name = Trim$(tokens(0))
timezone = Trim$(tokens(1))
notes = Trim$(tokens(2))
On Error GoTo Err

If name = "" Then
    gCon.writeErrorLine "Line " & lineNumber & ": name must be supplied"
    validInput = False
End If

If timezone = "" Then
    gCon.writeErrorLine "Line " & lineNumber & ": timezone must be supplied"
    validInput = False
End If

If Not validInput Then Exit Sub

If validInput Then
    Dim ex As exchange
    
    Set ex = gDb.ExchangeFactory.loadByName(name)
    If ex Is Nothing Then
        Set ex = gDb.ExchangeFactory.makeNew
    ElseIf Not gUpdate Then
        gCon.writeErrorLine "Line " & lineNumber & ": " & name & " already exists"
        Exit Sub
    Else
        update = True
    End If
    
    ex.name = name
    ex.TimeZoneName = timezone
    ex.notes = notes
    
    If ex.IsValid Then
        ex.ApplyEdit
        If update Then
            gCon.writeLineToConsole "Updated: " & name
        Else
            gCon.writeLineToConsole "Added: " & name
        End If
    Else
        Dim lErr As ErrorItem
        For Each lErr In ex.ErrorList
            Select Case lErr.ruleId
            Case BusinessRuleIds.BusRuleExchangeNameValid
                gCon.writeErrorLine "Line " & lineNumber & " name invalid"
            Case BusinessRuleIds.BusRuleExchangeTimezoneValid
                gCon.writeErrorLine "Line " & lineNumber & " timezone invalid"
            End Select
        Next
    End If
End If
        
Exit Sub

Err:
gCon.writeErrorLine Err.Description
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
    password = gCon.readLineFromConsole("Password:", "*")
End If
    
dbtype = DatabaseTypeFromString(dbtypeStr)
If dbtype = DbNone Then
    gCon.writeErrorLine "Error: invalid dbtype"
    setupDb = False
End If
    
If setupDb Then
    Set gDb = CreateTradingDB(GenerateConnectionString(dbtype, server, database, username, password))
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupDb = False

End Function

Private Sub showUsage()
gCon.writeErrorLine "Usage:"
gCon.writeErrorLine "uxd -todb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.writeErrorLine "    -U     # update existing records"
gCon.writeErrorLine "StdIn Formats:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "name,timezone[,notes]"
End Sub


