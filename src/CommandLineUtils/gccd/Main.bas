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

Private Const ProjectName                   As String = "gccd27"
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
If Not gCon Is Nothing Then gCon.WriteErrorLine Err.Description & " (" & Err.Source & ")"
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

Private Sub processContractClass( _
            ByVal cc As InstrumentClass)
gCon.WriteString cc.name
gCon.WriteString ","
gCon.WriteString cc.SecTypeString
gCon.WriteString ","
gCon.WriteString cc.CurrencyCode
gCon.WriteString ","
gCon.WriteString cc.TickSize
gCon.WriteString ","
gCon.WriteString cc.TickValue
gCon.WriteString ","
If cc.DaysBeforeExpiryToSwitch <> 0 Then gCon.WriteString cc.DaysBeforeExpiryToSwitch
gCon.WriteString ","
gCon.WriteString cc.SessionStartTime
gCon.WriteString ","
gCon.WriteString cc.SessionEndTime
gCon.WriteString ","
gCon.WriteLine cc.Notes
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
        Set summs = gDb.InstrumentClassFactory.Query(scb.conditionString, fieldnames)
        If summs.Count <> 0 Then
            gCon.WriteLine "$exchange " & exchange
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.LoadByID(summ.Id)
            Next
        End If
    Else
        Dim exSumms As DataObjectSummaries
        Dim exSumm As DataObjectSummary
        Set exSumms = gDb.ExchangeFactory.Query("", fieldnames)
        For Each exSumm In exSumms
            scb.Clear
            scb.addTerm "Exchange", CondOpEqual, exSumm.FieldValue("name")
            Set summs = gDb.InstrumentClassFactory.Query(scb.conditionString, fieldnames)
            gCon.WriteLine "$exchange " & exSumm.FieldValue("name")
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.LoadByID(summ.Id)
            Next
        Next
    End If
Else
    If exchange <> "" And exchange <> "*" Then
        scb.addTerm "name", CondOpEqual, name
        scb.addTerm "Exchange", CondOpEqual, exchange, LogicalOpAND
        Set summs = gDb.InstrumentClassFactory.Query(scb.conditionString, fieldnames)
        If summs.Count <> 0 Then
            gCon.WriteLine "$exchange " & exchange
            For Each summ In summs
                processContractClass gDb.InstrumentClassFactory.LoadByID(summ.Id)
            Next
        End If
    Else
        scb.addTerm "name", CondOpEqual, name
        Set summs = gDb.InstrumentClassFactory.Query(scb.conditionString, fieldnames)
        For Each summ In summs
            gCon.WriteLine "$exchange " & summ.FieldValue("exchange")
            processContractClass gDb.InstrumentClassFactory.LoadByID(summ.Id)
        Next
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
    Set gDb = CreateTradingDB(CreateConnectionParams(dbtype, _
                                                        server, _
                                                        database, _
                                                        username, _
                                                        password))
End If

Exit Function

Err:
gCon.WriteErrorLine Err.Description
setupDb = False

End Function

Private Sub showUsage()
gCon.WriteErrorLine "Usage:"
gCon.WriteErrorLine "gccd27 -fromdb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.WriteErrorLine ""
gCon.WriteErrorLine "StdIn Format:"
gCon.WriteErrorLine "#comment"
gCon.WriteErrorLine "$echo text"
gCon.WriteErrorLine "[* | name][,[exchange]]"
End Sub
