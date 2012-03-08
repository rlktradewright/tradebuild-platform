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

Public Const ProjectName                    As String = "ltz27"
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
    If clp.Switch("todb") Then
        If setupDb(clp.switchValue("todb")) Then
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
Dim s() As String
Dim tzf As TimeZoneFactory
Dim tz As TradingDO27.TimeZone
Dim i As Long

Set tzf = gDb.TimeZoneFactory

s = GetAvailableTimeZoneNames

For i = 0 To UBound(s)
    Set tz = tzf.LoadByName(s(i))
    If Not tz Is Nothing Then
        gCon.WriteLine s(i) & " already exists"
    Else
        Set tz = tzf.MakeNew
        tz.Name = s(i)
        tz.ApplyEdit
        gCon.WriteLine s(i) & " added"
    End If
Next


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
gCon.WriteErrorLine "ltz27 -todb:databaseserver,databasetype,catalog[,username[,password]]"
End Sub




