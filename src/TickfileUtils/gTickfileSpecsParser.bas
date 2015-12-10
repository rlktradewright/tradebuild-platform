Attribute VB_Name = "gTickfileSpecsParser"
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

Private Type State
    RootDir                                 As String
    From                                    As Date
    To                                      As Date
    SessionOnly                             As Boolean

    LineRootDir                             As String
    LineFrom                                As Date
    LineTo                                  As Date
    LineSessionOnly                         As Boolean
    LineSessionOnlyIsSet                    As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TickfileSpecsParser"

Private Const NoDate                                As Date = 0

Private Const FromSwitch                            As String = "from"
Private Const RootSwitch                            As String = "root"
Private Const ToSwitch                              As String = "to"
Private Const SessionOnlySwitch                     As String = "session"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function gParseTickfileListFile(ByVal pFilename As String) As TickFileSpecifiers
Const ProcName As String = "gParseTickfileListFile"
On Error GoTo Err

Dim lTickFileSpecifiers As New TickFileSpecifiers
Dim lFileSys As New Scripting.FileSystemObject

Dim lState As State

Dim lTs As Scripting.TextStream
Set lTs = lFileSys.OpenTextFile(pFilename, ForReading, False)

Do While Not lTs.AtEndOfStream
    processLine lTs.ReadLine, lState, lTickFileSpecifiers, lFileSys
Loop

lTs.Close

Set gParseTickfileListFile = lTickFileSpecifiers

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getDate( _
                ByVal pSwitchName As String, _
                ByVal pClp As CommandLineParser) As Date
Const ProcName As String = "getDate"
On Error GoTo Err

If pClp.SwitchValue(pSwitchName) = "" Then
    getDate = NoDate
Else
    AssertArgument IsDate(pClp.SwitchValue(pSwitchName)), "'" & pSwitchName & "' is not a valid date"
    getDate = pClp.SwitchValue(pSwitchName)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFromDate( _
                ByVal pClp As CommandLineParser) As Date
Const ProcName As String = "getFromDate"
On Error GoTo Err

getFromDate = getDate(FromSwitch, pClp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getRootDir( _
                ByVal pClp As CommandLineParser, _
                ByVal pFileSys As FileSystemObject) As String
Const ProcName As String = "getRootDir"
On Error GoTo Err

getRootDir = pClp.SwitchValue(RootSwitch)
AssertArgument pFileSys.FolderExists(getRootDir), "Specified root folder does not exist"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSessionOnly(ByVal pClp As CommandLineParser) As Boolean
Const ProcName As String = "getSessionOnly"
On Error GoTo Err

If LCase$(pClp.SwitchValue(SessionOnlySwitch)) = "off" Then
    getSessionOnly = False
ElseIf LCase$(pClp.SwitchValue(SessionOnlySwitch)) = "on" Then
    getSessionOnly = True
Else
    AssertArgument False, "'Session' switch value must be either 'on' or 'off'"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getToDate(ByVal pClp As CommandLineParser) As Date
Const ProcName As String = "getToDate"
On Error GoTo Err

getToDate = getDate(ToSwitch, pClp)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function processLine( _
                ByVal pLine As String, _
                ByRef pState As State, _
                ByVal pTickFileSpecifiers As TickFileSpecifiers, _
                ByVal pFileSys As FileSystemObject)
Const ProcName As String = "processLine"
On Error GoTo Err

Dim lClp As CommandLineParser
Set lClp = CreateCommandLineParser(pLine)

If lClp.NumberOfArgs = 0 Then
    If lClp.Switch(RootSwitch) Then pState.RootDir = getRootDir(lClp, pFileSys)
    If lClp.Switch(SessionOnlySwitch) Then pState.SessionOnly = getSessionOnly(lClp)
    If lClp.Switch(FromSwitch) Then pState.From = getFromDate(lClp)
    If lClp.Switch(ToSwitch) Then pState.To = getToDate(lClp)
    If pState.From <> NoDate And pState.To <> NoDate Then AssertArgument pState.From < pState.To, "From date is not earlier than To date"
Else
    pState.LineRootDir = ""
    pState.LineSessionOnlyIsSet = False
    pState.LineSessionOnly = False
    pState.LineFrom = NoDate
    pState.LineTo = NoDate

    If lClp.Switch(RootSwitch) Then pState.LineRootDir = getRootDir(lClp, pFileSys)
    If lClp.Switch(SessionOnlySwitch) Then
        pState.LineSessionOnlyIsSet = True
        pState.LineSessionOnly = getSessionOnly(lClp)
    End If
    If lClp.Switch(FromSwitch) Then pState.LineFrom = getFromDate(lClp)
    If lClp.Switch(ToSwitch) Then pState.LineTo = getToDate(lClp)
    If pState.LineFrom <> NoDate And pState.LineTo <> NoDate Then AssertArgument pState.LineFrom < pState.LineTo, "From date is not earlier than To date"
End If

Dim lRootDir As String
lRootDir = IIf(pState.LineRootDir <> "" And pState.LineRootDir <> "\", pState.LineRootDir, pState.RootDir)
AssertArgument lRootDir <> "" And lRootDir <> "\", "No root folder specified"
If Right$(lRootDir, 1) <> "\" Then lRootDir = lRootDir & "\"

Dim lTickfileName As String
Dim i As Long
For i = 0 To lClp.NumberOfArgs - 1
    pTickFileSpecifiers.Add processTickfileName( _
                                lClp.Arg(i), _
                                lRootDir, _
                                IIf(pState.LineSessionOnlyIsSet, pState.LineSessionOnly, pState.SessionOnly), _
                                IIf(pState.LineFrom <> NoDate, pState.LineFrom, pState.From), _
                                IIf(pState.LineTo <> NoDate, pState.LineTo, pState.To), _
                                pFileSys)
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function processTickfileName( _
                ByVal pTickfileName As String, _
                ByVal pRootDir As String, _
                ByVal pSessionOnly As Boolean, _
                ByVal pFromDate As Date, _
                ByVal pToDate As Date, _
                ByVal pFileSys As FileSystemObject) As TickfileSpecifier
Const ProcName As String = "processTickfileName"
On Error GoTo Err

Dim lSepPosn As Long
lSepPosn = InStr(1, pTickfileName, "-")
AssertArgument lSepPosn >= 2, "Invalid tickfile name: " & pTickfileName

Dim lSymbol As String
lSymbol = Left$(pTickfileName, lSepPosn - 1)

Dim lTickfileSpecifier As TickfileSpecifier
Set lTickfileSpecifier = New TickfileSpecifier

lTickfileSpecifier.Filename = pRootDir & lSymbol & "\" & pTickfileName
AssertArgument pFileSys.FileExists(lTickfileSpecifier.Filename), "Tickfile does not exist: " & lTickfileSpecifier.Filename

lTickfileSpecifier.EntireSession = pSessionOnly
lTickfileSpecifier.FromDate = pFromDate
lTickfileSpecifier.ToDate = pToDate

Set processTickfileName = lTickfileSpecifier

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


