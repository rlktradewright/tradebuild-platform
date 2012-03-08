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

Public Const ProjectName                    As String = "uccd27"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const ExchangeCommand               As String = "$EXCHANGE"

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

Public gDb As TradingDB
Public gExchange As Exchange

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
            If clp.Arg(0) = "" Then
                process
            Else
                Set gExchange = gDb.ExchangeFactory.LoadByName(clp.Arg(0))
                If gExchange Is Nothing Then
                    gCon.WriteErrorLine clp.Arg(0) & " is not a valid exchange"
                Else
                    process
                End If
            End If
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
    ElseIf Left$(inString, 1) = "$" Then
        ' process command
        If Len(inString) >= Len(ExchangeCommand) And _
            UCase$(Left$(inString, Len(ExchangeCommand))) = ExchangeCommand _
        Then
            Dim ex As String
            
            ex = Trim$(Right$(inString, Len(inString) - Len(ExchangeCommand)))
            gCon.WriteLineToConsole "Using exchange " & ex
            Set gExchange = gDb.ExchangeFactory.LoadByName(ex)
            If gExchange Is Nothing Then
                gCon.WriteErrorLine ex & " is not a valid exchange"
            End If
        End If
    Else
        If gExchange Is Nothing Then
            gCon.WriteErrorLine "No exchange defined"
            Exit Do
        End If
        processInput inString, lineNumber
    End If
    inString = Trim$(gCon.ReadLine(":"))
Loop
End Sub

Private Sub processInput( _
                ByVal inString As String, _
                ByVal lineNumber As Long)
' StdIn format:
' name,sectype,currency,ticksize,tickvalue,sessionstarttime,sessionendtime
Dim clp As CommandLineParser
Dim validInput As Boolean
Dim sectypeStr As String
Dim name As String
Dim currencyCode As String
Dim tickSizeStr As String
Dim tickValueStr As String
Dim switchday As String
Dim sessionStartStr As String
Dim sessionEndStr As String
Dim notes As String

Dim update As Boolean

Dim failpoint As Long
On Error GoTo Err

validInput = True

Set clp = CreateCommandLineParser(inString, InputSep)

name = clp.Arg(0)
sectypeStr = clp.Arg(1)
currencyCode = clp.Arg(2)
tickSizeStr = clp.Arg(3)
tickValueStr = clp.Arg(4)
switchday = clp.Arg(5)
sessionStartStr = clp.Arg(6)
sessionEndStr = clp.Arg(7)
notes = clp.Arg(8)

If name = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": name must be supplied"
    validInput = False
End If

If sectypeStr = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": sec type must be supplied"
    validInput = False
End If

If currencyCode = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": currency must be supplied"
    validInput = False
End If

If tickSizeStr = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": ticksize must be supplied"
    validInput = False
End If

If tickValueStr = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": tickvalue must be supplied"
    validInput = False
End If

If sessionStartStr = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": session start must be supplied"
    validInput = False
End If

If sessionEndStr = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": session end must be supplied"
    validInput = False
End If

If validInput Then
    Dim lInstrCl As InstrumentClass
    Dim contractSpec As ContractSpecifier
    Dim instrClassName As String
    
    instrClassName = gExchange.name & "/""" & name & """"
    Set lInstrCl = gDb.InstrumentClassFactory.LoadByName(instrClassName)
    If lInstrCl Is Nothing Then
        Set lInstrCl = gDb.InstrumentClassFactory.MakeNew
    ElseIf Not gUpdate Then
        gCon.WriteErrorLine "Line " & lineNumber & ": Already exists: " & instrClassName
        Exit Sub
    Else
        update = True
    End If
    
    lInstrCl.Exchange = gExchange
    lInstrCl.name = name
    lInstrCl.SecTypeString = sectypeStr
    lInstrCl.currencyCode = currencyCode
    lInstrCl.TickSizeString = tickSizeStr
    lInstrCl.TickValueString = tickValueStr
    lInstrCl.DaysBeforeExpiryToSwitchString = switchday
    lInstrCl.SessionStartTimeString = sessionStartStr
    lInstrCl.SessionEndTimeString = sessionEndStr
    lInstrCl.notes = notes
    
    If lInstrCl.IsValid Then
        lInstrCl.ApplyEdit
        If update Then
            gCon.WriteLineToConsole "Updated: " & instrClassName
        Else
            gCon.WriteLineToConsole "Added: " & instrClassName
        End If
    Else
        Dim lErr As ErrorItem
        For Each lErr In lInstrCl.ErrorList
            Select Case lErr.ruleId
            Case BusinessRuleIds.BusRuleInstrumentClassNameValid
                gCon.WriteErrorLine "Line " & lineNumber & " name invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassCurrencyCodeValid
                gCon.WriteErrorLine "Line " & lineNumber & " currency invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassDaysBeforeExpiryValid
                gCon.WriteErrorLine "Line " & lineNumber & " switchday invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassSecTypeValid
                gCon.WriteErrorLine "Line " & lineNumber & " sectype invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassSessionEndTimeValid
                gCon.WriteErrorLine "Line " & lineNumber & " sessionend invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassSessionEndTimeValid
                gCon.WriteErrorLine "Line " & lineNumber & " sessionstart invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassTickSizeValid
                gCon.WriteErrorLine "Line " & lineNumber & " ticksize invalid"
            Case BusinessRuleIds.BusRuleInstrumentClassTickValueValid
                gCon.WriteErrorLine "Line " & lineNumber & " tickvalue invalid"
            End Select
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
gCon.WriteErrorLine "uccd27 [exchange]"
gCon.WriteErrorLine "    -todb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.WriteErrorLine "    -U     # update existing records"
gCon.WriteErrorLine "StdIn Formats:"
gCon.WriteErrorLine "#comment"
gCon.WriteErrorLine "$class exchange"
gCon.WriteErrorLine "name,sectype,currency,ticksize,tickvalue,[switchday],sessionstarttime,sessionendtime[,notes]"
End Sub


