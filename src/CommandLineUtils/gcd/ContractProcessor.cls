VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements LogListener

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
Private Const ModuleName                    As String = "ContractProcessor"

Private Const OutputSep                     As String = ","

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec As ContractSpecifier
Private WithEvents mContractsTC As TaskController
Attribute mContractsTC.VB_VarHelpID = -1
Private WithEvents mTB As TradeBuildAPI
Attribute mTB.VB_VarHelpID = -1

Private mContracts As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mDone As Boolean
Private mLineNumber As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTB = TradeBuildAPI
GetLogger("log").AddLogListener Me
GetLogger("log").LogLevel = LogLevelNormal
End Sub

'@================================================================================
' InfoListener Interface Members
'@================================================================================

Private Sub LogListener_finish()

End Sub

Private Sub LogListener_Notify(ByVal logrec As TWUtilities30.LogRecord)
gCon.writeErrorLine VariantToString(logrec.Data)
End Sub

'@================================================================================
' mContracts Event Handlers
'@================================================================================

Private Sub mContracts_ContractSpecifierInvalid(ByVal reason As String)
gCon.writeErrorLine "Line " & mLineNumber & ": Invalid contract spec " & mContractSpec.ToString
mDone = True
End Sub

Private Sub mContracts_NoMoreContractDetails()
mDone = True
End Sub

Private Sub mContracts_Notification(ev As TWUtilities30.NotificationEvent)
gCon.writeErrorLine "Line " & mLineNumber & ": Notification: " & ev.eventMessage
mDone = True
End Sub

Private Sub mContracts_TryLater(ByVal reason As String, ByVal serviceProviderName As String)
gCon.writeErrorLine "Line " & mLineNumber & ": Try later: " & reason
mDone = True
End Sub

Private Sub mContractsTC_Completed(ev As TWUtilities30.TaskCompletionEvent)
If ev.cancelled Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Contract fetch cancelled"
ElseIf ev.errorNumber <> 0 Then
    gCon.writeErrorLine "Line " & mLineNumber & ": Error " & ev.errorNumber & ": " & ev.errorMessage
Else
    If Not IsEmpty(ev.result) Then Set mContracts = ev.result
End If
mDone = True
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Error(ev As TWUtilities30.ErrorEvent)
gCon.writeErrorLine "TradeBuild error: " & ev.errorMessage
End Sub

Private Sub mTB_Notification(ev As TWUtilities30.NotificationEvent)
Dim spe As ServiceProviderError
If ev.eventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spe = mTB.GetServiceProviderError
    gCon.writeErrorLine "Service provider error: " & spe.message
Else
    gCon.writeErrorLine "TradeBuild notification: " & ev.eventMessage
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub process( _
                ByVal contractSpec As ContractSpecifier, _
                ByVal lineNumber As Long, _
                ByVal nametemplate As String)
Dim lContract As Contract

mDone = False

Set mContracts = Nothing
Set mContractSpec = contractSpec
mLineNumber = lineNumber
Set mContractsTC = TradeBuildAPI.loadContracts(contractSpec)

' now hang around until we've got the info, but don't exit this proc
' because we don't want to process the next line of input till this one's done

Do While Not mDone
    Wait (100)
Loop

If mContracts.Count <> 0 Then
    For Each lContract In mContracts
        gCon.writeLine formatName(lContract, nametemplate) & OutputSep & _
                        lContract.specifier.localSymbol & OutputSep & _
                        lContract.specifier.symbol & OutputSep & _
                        lContract.specifier.expiry & OutputSep & _
                        lContract.specifier.strike & OutputSep & _
                        OptionRightToString(lContract.specifier.Right) & OutputSep & _
                        SecTypeToString(lContract.specifier.sectype) & OutputSep & _
                        lContract.specifier.exchange & OutputSep & _
                        lContract.specifier.currencyCode & OutputSep & _
                        lContract.TickSize & OutputSep & _
                        lContract.tickValue & OutputSep & _
                        Format(lContract.sessionStartTime, "hh:mm") & OutputSep & _
                        Format(lContract.sessionEndTime, "hh:mm") & OutputSep & _
                        lContract.timezone.standardName & OutputSep & _
                        lContract.daysBeforeExpiryToSwitch & OutputSep & _
                        lContract.numberOfDecimals & OutputSep & _
                        lContract.priceFormatString & OutputSep & _
                        lContract.multiplier
    Next
End If

End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================

Private Function formatName( _
                ByVal pContract As Contract, _
                ByVal nametemplate As String) As String
Dim curr As String
Dim s As StringStreamer
Dim bldr As StringBuilder

If nametemplate = "" Then
    formatName = pContract.Description
    Exit Function
End If

Set s = New StringStreamer
s.initialise nametemplate

Set bldr = CreateStringBuilder

bldr.Append """"
Do While Not s.eos
    curr = s.nextChar
    If curr = "\" Then
        bldr.Append s.nextChar
    ElseIf curr = "%" Then
        curr = s.nextChar
        Select Case curr
        Case "d"          'date
            Dim formatString As String
            curr = s.peekChar
            If curr <> "{" Then
                formatString = "ddddd"
            Else
                s.nextChar
                Do While Not s.eos
                    curr = s.nextChar
                    If curr = "}" Then
                        Exit Do
                    Else
                        formatString = formatString & curr
                    End If
                Loop
            End If
            If Not pContract.expiryDate = 0 Then bldr.Append Format(pContract.expiryDate, formatString)
        Case "s"          'symbol
            bldr.Append pContract.specifier.symbol
        Case "S"          'shortname"
            bldr.Append pContract.specifier.localSymbol
        Case "e"          'exchange"
            bldr.Append pContract.specifier.exchange
        Case "c"          'currency"
            bldr.Append pContract.specifier.currencyCode
        Case "t"          'sectype (abbreviated)"
            bldr.Append SecTypeToShortString(pContract.specifier.sectype)
        Case "T"          'sectype (in full)"
            bldr.Append SecTypeToString(pContract.specifier.sectype)
        Case "p"          'strike price"
            If pContract.specifier.strike <> 0 Then bldr.Append pContract.specifier.strike
        Case "r"          'right"
            bldr.Append OptionRightToString(pContract.specifier.Right)
        Case "n"          'supplied name"
            bldr.Append pContract.Description
        Case Else
            bldr.Append "%" & curr
        End Select
    Else
        bldr.Append curr
    End If
Loop
bldr.Append """"

formatName = bldr.ToString

End Function


