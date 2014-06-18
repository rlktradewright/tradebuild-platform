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

Public Const ProjectName                    As String = "ucd27"
Private Const ModuleName                    As String = "MainMod"

Private Const InputSep                      As String = ","

Private Const ClassCommand                   As String = "$CLASS"

'@================================================================================
' Member variables
'@================================================================================

Public gCon As Console

Public gDb As TradingDB
Public gInstrumentClass As InstrumentClass

' if set, existing records are to be updated
Public gUpdate As Boolean

Public gAllowOverrides  As Boolean

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
            If clp.Switch("O") Then gAllowOverrides = True
            If clp.Arg(0) = "" Then
                process
            Else
                Set gInstrumentClass = gDb.InstrumentClassFactory.LoadByName(clp.Arg(0))
                If gInstrumentClass Is Nothing Then
                    gCon.WriteErrorLine clp.Arg(0) & " is not a valid contract class"
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
        If Len(inString) >= Len(ClassCommand) And _
            UCase$(Left$(inString, Len(ClassCommand))) = ClassCommand _
        Then
            Dim class As String
            
            class = Trim$(Right$(inString, Len(inString) - Len(ClassCommand)))
            gCon.WriteLineToConsole "Using contract class " & class
            Set gInstrumentClass = gDb.InstrumentClassFactory.LoadByName(class)
            If gInstrumentClass Is Nothing Then
                gCon.WriteErrorLine class & " is not a valid contract class"
            End If
        End If
    Else
        If gInstrumentClass Is Nothing Then
            gCon.WriteErrorLine "No contract class defined"
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
' name,shortname,symbol,expiry,strike,right[,[sectype][,[exchange][,[currency][,[ticksize][,[tickvalue]]]]]]
Dim validInput As Boolean

Dim sectype As SecurityTypes
Dim sectypeStr As String
Dim exchange As String
Dim name As String
Dim shortname As String
Dim symbol As String
Dim currencyCode As String
Dim expiry As String
Dim expiryDate As Date
Dim strike As Double
Dim strikeStr As String
Dim optRight As OptionRights
Dim optRightStr As String
Dim tickSizeStr As String
Dim tickSize As Double
Dim tickValueStr As String
Dim tickValue As Double
Dim update As Boolean

Dim parser As CommandLineParser

Dim failpoint As Long
On Error GoTo Err

validInput = True

Set parser = CreateCommandLineParser(inString, ",")

name = parser.Arg(0)
shortname = parser.Arg(1)
symbol = parser.Arg(2)
expiry = parser.Arg(3)
strikeStr = parser.Arg(4)
optRightStr = parser.Arg(5)
sectypeStr = parser.Arg(6)
exchange = parser.Arg(7)
currencyCode = parser.Arg(8)
tickSizeStr = parser.Arg(9)
tickValueStr = parser.Arg(10)

If name = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": name must be supplied"
    validInput = False
End If

If shortname = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": shortname must be supplied"
    validInput = False
End If

If symbol = "" Then
    gCon.WriteErrorLine "Line " & lineNumber & ": symbol must be supplied"
    validInput = False
End If

If gInstrumentClass.sectype = SecTypeFuture Or _
    gInstrumentClass.sectype = SecTypeOption Or _
    gInstrumentClass.sectype = SecTypeFuturesOption _
Then
    If expiry = "" Then
        gCon.WriteErrorLine "Line " & lineNumber & ": expiry must be supplied"
        validInput = False
    End If
End If

If gInstrumentClass.sectype = SecTypeOption Or _
    gInstrumentClass.sectype = SecTypeFuturesOption _
Then
    If strike = 0 Then
        gCon.WriteErrorLine "Line " & lineNumber & ": strike must be supplied"
        validInput = False
    End If
    If optRight = OptNone Then
        gCon.WriteErrorLine "Line " & lineNumber & ": right must be supplied"
        validInput = False
    End If
End If

sectype = SecTypeFromString(sectypeStr)
If sectypeStr <> "" And sectype = SecTypeNone Then
    gCon.WriteErrorLine "Line " & lineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validInput = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiryDate = CDate(expiry)
    ElseIf Len(expiry) = 8 Then
        If IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            expiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
        Else
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

If tickSizeStr <> "" Then
    If IsNumeric(tickSizeStr) Then
        tickSize = CDbl(tickSizeStr)
    Else
        gCon.WriteErrorLine "Line " & lineNumber & ": Invalid ticksize '" & tickSizeStr & "'"
        validInput = False
    End If
End If

If tickValueStr <> "" Then
    If IsNumeric(tickValueStr) Then
        tickValue = CDbl(tickValueStr)
    Else
        gCon.WriteErrorLine "Line " & lineNumber & ": Invalid tickvalue '" & tickValueStr & "'"
        validInput = False
    End If
End If

optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.WriteErrorLine "Line " & lineNumber & ": Invalid right '" & optRightStr & "'"
    validInput = False
End If

If Not validInput Then Exit Sub

If (exchange <> "" And exchange <> gInstrumentClass.ExchangeName) Or _
    (sectype <> SecTypeNone And sectype <> gInstrumentClass.sectype) Or _
    (Not gAllowOverrides And _
        ((currencyCode <> "" And currencyCode <> gInstrumentClass.currencyCode) Or _
        (tickSize <> 0 And tickSize <> gInstrumentClass.tickSize) Or _
        (tickValue <> 0 And tickValue <> gInstrumentClass.tickValue)) _
    ) _
Then
    gCon.WriteErrorLine "Line " & lineNumber & ": details do not match contract class " & gInstrumentClass.exchange.name & "/" & gInstrumentClass.name & ":"
    gCon.WriteErrorLine CreateContractSpecifier( _
                                                shortname, _
                                                symbol, _
                                                exchange, _
                                                sectype, _
                                                currencyCode, _
                                                expiry, _
                                                strike, _
                                                optRight).ToString & _
                        "; tickSize=" & tickSize & _
                        "; tickValue=" & tickValue
    validInput = False
End If

If Not validInput Then
    gCon.WriteErrorLine inString
Else
    'Dim lInstr As Instrument
    'Dim scb As New SimpleConditionBuilder
    '
    'scb.addTerm "shortname", CondOpEqual, shortname
    'Set lInstr = gDb.InstrumentFactory.LoadByQuery(scb.conditionString)
    
    Dim lInstr As Instrument
    Set lInstr = gDb.InstrumentFactory.LoadByName(gInstrumentClass.ExchangeName & "/" & gInstrumentClass.name & "/" & name)
    If lInstr Is Nothing Then
        Set lInstr = gDb.InstrumentFactory.MakeNew
    ElseIf Not gUpdate Then
        gCon.WriteErrorLine "Line " & lineNumber & ": Already exists: " & gInstrumentClass.ExchangeName & "/" & gInstrumentClass.name & "/" & name & "(" & shortname & ")"
        Exit Sub
    Else
        update = True
    End If
    
    lInstr.InstrumentClass = gInstrumentClass
    lInstr.name = name
    lInstr.shortname = shortname
    lInstr.symbol = symbol
    If gInstrumentClass.sectype = SecTypeFuture Then
        lInstr.expiryDate = expiryDate
    ElseIf gInstrumentClass.sectype = SecTypeOption Or _
        gInstrumentClass.sectype = SecTypeFuturesOption _
    Then
        lInstr.expiryDate = expiryDate
        lInstr.StrikePrice = strike
        lInstr.OptionRight = optRight
    End If
    
    If tickSize <> 0 And tickSize <> gInstrumentClass.tickSize Then
        lInstr.tickSize = tickSize
    Else
        lInstr.TickSizeString = ""
    End If
    
    If tickValue <> 0 And tickValue <> gInstrumentClass.tickValue Then
        lInstr.tickValue = tickValue
    Else
        lInstr.TickValueString = ""
    End If
    
    If currencyCode <> "" And currencyCode <> gInstrumentClass.currencyCode Then
        lInstr.currencyCode = currencyCode
    Else
        lInstr.currencyCode = ""
    End If
    
    If lInstr.IsValid Then
        lInstr.ApplyEdit
        If update Then
            gCon.WriteLineToConsole "Updated: " & gInstrumentClass.ExchangeName & "/" & gInstrumentClass.name & "/" & name & " (" & shortname & ")"
        Else
            gCon.WriteLineToConsole "Added:   " & gInstrumentClass.ExchangeName & "/" & gInstrumentClass.name & "/" & name & " (" & shortname & ")"
        End If
    Else
        Dim lErr As ErrorItem
        For Each lErr In lInstr.ErrorList
            Select Case lErr.RuleId
            Case BusinessRuleIds.BusRuleInstrumentNameValid
                gCon.WriteErrorLine "Line " & lineNumber & ": name invalid: " & "" & name & ""
            Case BusinessRuleIds.BusRuleInstrumentOptionRightvalid
                gCon.WriteErrorLine "Line " & lineNumber & ": right invalid"
            Case BusinessRuleIds.BusRuleInstrumentShortNameValid
                gCon.WriteErrorLine "Line " & lineNumber & ": shortname invalid"
            Case BusinessRuleIds.BusRuleInstrumentStrikePriceValid
                gCon.WriteErrorLine "Line " & lineNumber & ": strike invalid"
            Case BusinessRuleIds.BusRuleInstrumentSymbolValid
                gCon.WriteErrorLine "Line " & lineNumber & ": symbol invalid"
            Case BusinessRuleIds.BusRuleInstrumentTickSizeValid
                gCon.WriteErrorLine "Line " & lineNumber & ": ticksize invalid"
            Case BusinessRuleIds.BusRuleInstrumentTickValueValid
                gCon.WriteErrorLine "Line; " & lineNumber & ": tickvalue invalid"
            End Select
        Next
        gCon.WriteErrorLine inString
    End If
End If
        
Exit Sub

Err:
gCon.WriteErrorLine Err.Description & " when processing:"
gCon.WriteErrorLine inString
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
gCon.WriteErrorLine "ucd27 [exchange/contractclass]"
gCon.WriteErrorLine "    -todb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.WriteErrorLine "    -U     # update existing records"
gCon.WriteErrorLine "    -O     # allow overrides to Contract Class ticksize, tickvalue and currency"
gCon.WriteErrorLine "StdIn Formats:"
gCon.WriteErrorLine "#comment"
gCon.WriteErrorLine "$class exchange/contractclass"
gCon.WriteErrorLine "name,shortname,symbol,expiry,strike,right[,[sectype][,[exchange][,[currency][,[ticksize][,[tickvalue]]]]]]"
End Sub


