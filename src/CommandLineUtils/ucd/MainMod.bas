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

Private Const ProjectName                   As String = "ucd"
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
                Set gInstrumentClass = gDb.InstrumentClassFactory.loadByName(clp.Arg(0))
                If gInstrumentClass Is Nothing Then
                    gCon.writeErrorLine clp.Arg(0) & " is not a valid contract class"
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
    ElseIf Left$(inString, 1) = "$" Then
        ' process command
        If Len(inString) >= Len(ClassCommand) And _
            UCase$(Left$(inString, Len(ClassCommand))) = ClassCommand _
        Then
            Dim class As String
            
            class = Trim$(Right$(inString, Len(inString) - Len(ClassCommand)))
            gCon.writeLineToConsole "Using contract class " & class
            Set gInstrumentClass = gDb.InstrumentClassFactory.loadByName(class)
            If gInstrumentClass Is Nothing Then
                gCon.writeErrorLine class & " is not a valid contract class"
            End If
        End If
    Else
        If gInstrumentClass Is Nothing Then
            gCon.writeErrorLine "No contract class defined"
            Exit Do
        End If
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

Dim failpoint As Long
On Error GoTo Err

validInput = True

tokens = Split(inString, InputSep)

On Error Resume Next
name = Trim$(tokens(0))
shortname = Trim$(tokens(1))
symbol = Trim$(tokens(2))
expiry = Trim$(tokens(3))
strikeStr = Trim$(tokens(4))
optRightStr = Trim$(tokens(5))
sectypeStr = Trim$(tokens(6))
exchange = Trim$(tokens(7))
currencyCode = Trim$(tokens(8))
tickSizeStr = Trim$(tokens(9))
tickValueStr = Trim$(tokens(10))
On Error GoTo Err

If name = "" Then
    gCon.writeErrorLine "Line " & lineNumber & ": name must be supplied"
    validInput = False
End If

If shortname = "" Then
    gCon.writeErrorLine "Line " & lineNumber & ": shortname must be supplied"
    validInput = False
End If

If symbol = "" Then
    gCon.writeErrorLine "Line " & lineNumber & ": symbol must be supplied"
    validInput = False
End If

If gInstrumentClass.sectype = SecTypeFuture Or _
    gInstrumentClass.sectype = SecTypeOption Or _
    gInstrumentClass.sectype = SecTypeFuturesOption _
Then
    If expiry = "" Then
        gCon.writeErrorLine "Line " & lineNumber & ": expiry must be supplied"
        validInput = False
    End If
End If

If gInstrumentClass.sectype = SecTypeOption Or _
    gInstrumentClass.sectype = SecTypeFuturesOption _
Then
    If strike = 0 Then
        gCon.writeErrorLine "Line " & lineNumber & ": strike must be supplied"
        validInput = False
    End If
    If optRight = OptNone Then
        gCon.writeErrorLine "Line " & lineNumber & ": right must be supplied"
        validInput = False
    End If
End If

sectype = SecTypeFromString(sectypeStr)
If sectypeStr <> "" And sectype = SecTypeNone Then
    gCon.writeErrorLine "Line " & lineNumber & ": Invalid sectype '" & sectypeStr & "'"
    validInput = False
End If

If expiry <> "" Then
    If IsDate(expiry) Then
        expiryDate = CDate(expiry)
    ElseIf Len(expiry) = 8 Then
        If IsDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2)) Then
            expiryDate = CDate(Left$(expiry, 4) & "/" & Mid$(expiry, 5, 2) & "/" & Right$(expiry, 2))
        Else
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

If tickSizeStr <> "" Then
    If IsNumeric(tickSizeStr) Then
        tickSize = CDbl(tickSizeStr)
    Else
        gCon.writeErrorLine "Line " & lineNumber & ": Invalid ticksize '" & tickSizeStr & "'"
        validInput = False
    End If
End If

If tickValueStr <> "" Then
    If IsNumeric(tickValueStr) Then
        tickValue = CDbl(tickValueStr)
    Else
        gCon.writeErrorLine "Line " & lineNumber & ": Invalid tickvalue '" & tickValueStr & "'"
        validInput = False
    End If
End If

optRight = OptionRightFromString(optRightStr)
If optRightStr <> "" And optRight = OptNone Then
    gCon.writeErrorLine "Line " & lineNumber & ": Invalid right '" & optRightStr & "'"
    validInput = False
End If

If Not validInput Then Exit Sub

If (exchange <> "" And exchange <> gInstrumentClass.exchangeName) Or _
    (sectype <> SecTypeNone And sectype <> gInstrumentClass.sectype) Or _
    (Not gAllowOverrides And _
        ((currencyCode <> "" And currencyCode <> gInstrumentClass.currencyCode) Or _
        (tickSize <> 0 And tickSize <> gInstrumentClass.tickSize) Or _
        (tickValue <> 0 And tickValue <> gInstrumentClass.tickValue)) _
    ) _
Then
    gCon.writeErrorLine "Line " & lineNumber & ": details do not match contract class " & gInstrumentClass.exchange.name & "/" & gInstrumentClass.name & ":"
    gCon.writeErrorLine CreateContractSpecifier( _
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

If validInput Then
    Dim lInstr As Instrument
    Dim scb As New SimpleConditionBuilder

    scb.addTerm "shortname", CondOpEqual, shortname
    Set lInstr = gDb.InstrumentFactory.loadByQuery(scb.conditionString)
    If lInstr Is Nothing Then
        Set lInstr = gDb.InstrumentFactory.makeNew
    ElseIf Not gUpdate Then
        gCon.writeErrorLine "Line " & lineNumber & ": Already exists: " & gInstrumentClass.exchangeName & "/" & gInstrumentClass.name & "/" & name & "(" & shortname & ")"
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
        lInstr.strikePrice = strike
        lInstr.optionRight = optRight
    End If
    
    If tickSize <> 0 And tickSize <> gInstrumentClass.tickSize Then
        lInstr.tickSize = tickSize
    Else
        lInstr.tickSizeString = ""
    End If
    
    If tickValue <> 0 And tickValue <> gInstrumentClass.tickValue Then
        lInstr.tickValue = tickValue
    Else
        lInstr.tickValueString = ""
    End If
    
    If currencyCode <> "" And currencyCode <> gInstrumentClass.currencyCode Then
        lInstr.currencyCode = currencyCode
    Else
        lInstr.currencyCode = ""
    End If
    
    If lInstr.IsValid Then
        lInstr.ApplyEdit
        If update Then
            gCon.writeLineToConsole "Updated: " & gInstrumentClass.exchangeName & "/" & gInstrumentClass.name & "/" & name & "(" & shortname & ")"
        Else
            gCon.writeLineToConsole "Added: " & gInstrumentClass.exchangeName & "/" & gInstrumentClass.name & "/" & name & "(" & shortname & ")"
        End If
    Else
        Dim lErr As ErrorItem
        For Each lErr In lInstr.ErrorList
            Select Case lErr.ruleId
            Case BusinessRuleIds.BusRuleInstrumentNameValid
                gCon.writeErrorLine "Line " & lineNumber & " name invalid"
            Case BusinessRuleIds.BusRuleInstrumentOptionRightvalid
                gCon.writeErrorLine "Line " & lineNumber & " right invalid"
            Case BusinessRuleIds.BusRuleInstrumentShortNameValid
                gCon.writeErrorLine "Line " & lineNumber & " shortname invalid"
            Case BusinessRuleIds.BusRuleInstrumentStrikePriceValid
                gCon.writeErrorLine "Line " & lineNumber & " strike invalid"
            Case BusinessRuleIds.BusRuleInstrumentSymbolValid
                gCon.writeErrorLine "Line " & lineNumber & " symbol invalid"
            Case BusinessRuleIds.BusRuleInstrumentTickSizeValid
                gCon.writeErrorLine "Line " & lineNumber & " ticksize invalid"
            Case BusinessRuleIds.BusRuleInstrumentTickValueValid
                gCon.writeErrorLine "Line " & lineNumber & " tickvalue invalid"
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
gCon.writeErrorLine "ucd [exchange/contractclass]"
gCon.writeErrorLine "    -todb:<databaseserver>,<databasetype>,<catalog>[,<username>[,<password>]]"
gCon.writeErrorLine "    -U     # update existing records"
gCon.writeErrorLine "    -O     # allow overrides to Contract Class ticksize, tickvalue and currency"
gCon.writeErrorLine "StdIn Formats:"
gCon.writeErrorLine "#comment"
gCon.writeErrorLine "$class exchange/contractclass"
gCon.writeErrorLine "name,shortname,symbol,expiry,strike,right[,[sectype][,[exchange][,[currency][,[ticksize][,[tickvalue]]]]]]"
End Sub


