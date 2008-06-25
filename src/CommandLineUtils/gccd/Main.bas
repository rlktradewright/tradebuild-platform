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

Private Const ProjectName                   As String = "gccd"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const EchoCommand                   As String = "$ECHO"

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
    clp.NumberOfArgs > 0 Or _
    clp.NumberOfSwitches = 0 _
Then
    showUsage
ElseIf clp.Switch("fromdb") Then
    If setupDb(clp.switchValue("fromdb")) Then
        process
    End If
Else
    showUsage
End If

TerminateTWUtilities

Exit Sub

Err:
If Not gCon Is Nothing Then gCon.writeErrorLine Err.Description & " (" & Err.Source & ")"
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

Private Sub processContractClass( _
            ByVal cc As InstrumentClass)
gCon.writeString cc.name
gCon.writeString ","
gCon.writeString cc.secTypeString
gCon.writeString ","
gCon.writeString cc.currencyCode
gCon.writeString ","
gCon.writeString cc.tickSize
gCon.writeString ","
gCon.writeString cc.tickValue
gCon.writeString ","
If cc.daysBeforeExpiryToSwitch <> 0 Then gCon.writeString cc.daysBeforeExpiryToSwitch
gCon.writeString ","
gCon.writeString cc.sessionStartTime
gCon.writeString ","
gCon.writeString cc.sessionEndTime
gCon.writeString ","
gCon.writeLine cc.notes
End Sub

Private Sub processInput( _
                ByVal inString As String, _
                ByVal lineNumber As Long)
' StdIn format:
' sectype,exchange,shortname,symbol,currency,expiry,strike,right,nametemplate

Dim tokens() As String
Dim name As String
Dim exchange As String
Dim scb As New SimpleConditionBuilder
Dim summs As DataObjectSummaries
Dim summ As DataObjectSummary
Dim fieldnames() As String

Dim failpoint As Long
On Error GoTo Err

tokens = Split(inString, InputSep)

On Error Resume Next
name = Trim$(tokens(0))
exchange = Trim$(tokens(1))
On Error GoTo Err

If name = "*" Then
    If exchange <> "" And exchange <> "*" Then
        scb.addTerm "Exchange", CondOpEqual, exchange
        Set summs = gDb.InstrumentClassFactory.query(scb.conditionString, fieldnames)
        If summs.Count <> 0 Then
            gCon.writeLine "$exchange " & exchange
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.loadByID(summ.id)
            Next
        End If
    Else
        Dim exSumms As DataObjectSummaries
        Dim exSumm As DataObjectSummary
        Set exSumms = gDb.ExchangeFactory.query("", fieldnames)
        For Each exSumm In exSumms
            scb.Clear
            scb.addTerm "Exchange", CondOpEqual, exSumm.fieldValue("name")
            Set summs = gDb.InstrumentClassFactory.query(scb.conditionString, fieldnames)
            gCon.writeLine "$exchange " & exSumm.fieldValue("name")
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.loadByID(summ.id)
            Next
        Next
    End If
Else
    If exchange <> "" And exchange <> "*" Then
        scb.addTerm "name", CondOpEqual, name
        scb.addTerm "Exchange", CondOpEqual, exchange, LogicalOpAND
        Set summs = gDb.InstrumentClassFactory.query(scb.conditionString, fieldnames)
        If summs.Count <> 0 Then
            gCon.writeLine "$exchange " & exchange
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.loadByID(summ.id)
            Next
        End If
    Else
        scb.addTerm "name", CondOpEqual, name
        Set summs = gDb.InstrumentClassFactory.query(scb.conditionString, fieldnames)
        For Each summ In summs
            gCon.writeLine "$exchange " & summ.fieldValue("exchange")
            processContractClass gDb.InstrumentClassFactory.loadByID(summ.id)
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
    Set gDb = CreateTradingDB(GenerateConnectionString(dbtype, _
                                                        server, _
                                                        database, _
                                                        username, _
                                                        password), _
                            dbtype)
End If

Exit Function

Err:
gCon.writeErrorLine Err.Description
setupDb = False

End Function

Private Sub showUsage()
gCon.writeErrorLine "Usage:"
gCon.writeErrorLine "gccd -fromdb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.writeErrorLine ""
gCon.writeErrorLine "StdIn Format:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "$echo text"
gCon.writeErrorLine "[* | name][,[exchange]]"
End Sub
