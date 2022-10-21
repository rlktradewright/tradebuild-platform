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

Implements ILogListener

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

Private Const ModuleName                    As String = "ContractProcessor"

Private Const OutputSep                     As String = ","

'@================================================================================
' Member variables
'@================================================================================

Private mContractSpec As ContractSpecifier
Private WithEvents mTB As TradeBuildAPI
Attribute mTB.VB_VarHelpID = -1

Private mContracts As Contracts
Attribute mContracts.VB_VarHelpID = -1

Private mLineNumber As Long

Private mNametemplate As String

Private WithEvents mFutureWaiter As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
GetLogger("log").AddLogListener Me
GetLogger("log").LogLevel = LogLevelNormal
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' InfoListener Interface Members
'@================================================================================

Private Sub ILogListener_finish()

End Sub

Private Sub ILogListener_Notify(ByVal logrec As LogRecord)
gCon.WriteErrorLine VariantToString(logrec.Data)
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If ev.Future.IsCancelled Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": contract fetch cancelled: " & mContractSpec.ToString
ElseIf ev.Future.IsFaulted Then
    gCon.WriteErrorLine "Line " & mLineNumber & ": contract fetch error: " & mContractSpec.ToString
Else
    Set mContracts = ev.Future.Value
    If mContracts.Count = 0 Then gCon.WriteErrorLine "Line " & mLineNumber & ": no contracts: " & mContractSpec.ToString
    processContracts
End If
gPauseInput = False
End Sub

'@================================================================================
' mTB Event Handlers
'@================================================================================

Private Sub mTB_Notification(ev As NotificationEventData)
Dim spe As ServiceProviderError
If ev.EventCode = ApiNotifyCodes.ApiNotifyServiceProviderError Then
    Set spe = mTB.GetServiceProviderError
    gCon.WriteErrorLine "Service provider error: " & spe.Message
Else
    gCon.WriteErrorLine "TradeBuild notification: " & ev.EventMessage
End If
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub initialise(ByVal pTB As TradeBuildAPI)
Set mTB = pTB
End Sub

Public Sub process( _
                ByVal contractSpec As IContractSpecifier, _
                ByVal lineNumber As Long, _
                ByVal nametemplate As String)
Set mContracts = Nothing
Set mContractSpec = contractSpec
mLineNumber = lineNumber
mNametemplate = nametemplate
mFutureWaiter.Add mTB.ContractStorePrimary.FetchContracts(contractSpec)
gPauseInput = True
End Sub

Private Sub processContracts()
If mContracts.Count = 0 Then Exit Sub

Dim lContract As IContract
For Each lContract In mContracts
    gCon.WriteLine formatName(lContract, mNametemplate) & OutputSep & _
                    lContract.Specifier.LocalSymbol & OutputSep & _
                    lContract.Specifier.symbol & OutputSep & _
                    lContract.Specifier.expiry & OutputSep & _
                    lContract.Specifier.multiplier & OutputSep & _
                    lContract.Specifier.strike & OutputSep & _
                    OptionRightToString(lContract.Specifier.Right) & OutputSep & _
                    SecTypeToString(lContract.Specifier.sectype) & OutputSep & _
                    lContract.Specifier.exchange & OutputSep & _
                    lContract.Specifier.currencyCode & OutputSep & _
                    lContract.TickSize & OutputSep & _
                    lContract.TickValue & OutputSep & _
                    Format(lContract.SessionStartTime, "hh:mm") & OutputSep & _
                    Format(lContract.SessionEndTime, "hh:mm") & OutputSep & _
                    lContract.TimezoneName & OutputSep & _
                    lContract.DaysBeforeExpiryToSwitch & OutputSep & _
                    lContract.NumberOfDecimals & OutputSep
Next

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
            If Not pContract.ExpiryDate = 0 Then bldr.Append Format(pContract.ExpiryDate, formatString)
        Case "s"          'symbol
            bldr.Append pContract.Specifier.symbol
        Case "S"          'shortname"
            bldr.Append pContract.Specifier.LocalSymbol
        Case "e"          'exchange"
            bldr.Append pContract.Specifier.exchange
        Case "c"          'currency"
            bldr.Append pContract.Specifier.currencyCode
        Case "t"          'sectype (abbreviated)"
            bldr.Append SecTypeToShortString(pContract.Specifier.sectype)
        Case "T"          'sectype (in full)"
            bldr.Append SecTypeToString(pContract.Specifier.sectype)
        Case "p"          'strike price"
            If pContract.Specifier.strike <> 0 Then bldr.Append pContract.Specifier.strike
        Case "r"          'right"
            bldr.Append OptionRightToString(pContract.Specifier.Right)
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


