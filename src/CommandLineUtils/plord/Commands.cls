VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Commands"
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

Private Const ModuleName                            As String = "Commands"

Private Const BatchOrdersCommandName                As String = "BATCHORDERS"
Private Const BracketCommandName                    As String = "BRACKET"
Private Const BuyCommandName                        As String = "BUY"
Private Const BuyAgainCommandName                   As String = "B"
Private Const CloseoutCommandName                   As String = "CLOSEOUT"
Private Const ContractCommandName                   As String = "CONTRACT"
Private Const EndBracketCommandName                 As String = "ENDBRACKET"
Private Const EndOrdersCommandName                  As String = "ENDORDERS"
Private Const EntryCommandName                      As String = "ENTRY"
Private Const ExitCommandName                       As String = "EXIT"
Private Const GroupCommandName                      As String = "GROUP"
Private Const HelpCommandName                       As String = "HELP"
Private Const Help1CommandName                      As String = "?"
Private Const ListCommandName                       As String = "LIST"
Private Const PurgeCommandName                      As String = "PURGE"
Private Const QuitCommandName                       As String = "QUIT"
Private Const QuoteCommandName                      As String = "QUOTE"
Private Const ResetCommandName                      As String = "RESET"
Private Const RolloverCommandName                   As String = "ROLLOVER"
Private Const SellCommandName                       As String = "SELL"
Private Const SellAgainCommandName                  As String = "S"
Private Const SetBalanceCommandName                 As String = "SETBALANCE"
Private Const SetFundsCommandName                   As String = "SETFUNDS"
Private Const SetGroupFundsCommandName              As String = "SETGROUPFUNDS"
Private Const SetGroupRolloverCommandName           As String = "SETGROUPROLLOVER"
Private Const SetRolloverCommandName                As String = "SETROLLOVER"
Private Const ShowBalanceCommandName                As String = "SHOWBALANCE"
Private Const StageOrdersCommandName                As String = "STAGEORDERS"
Private Const StopLossCommandName                   As String = "STOPLOSS"
Private Const TargetCommandName                     As String = "TARGET"

'@================================================================================
' Member variables
'@================================================================================

Private mBatchOrdersCommand                         As New Command
Private mBracketCommand                             As New Command
Private mBuyCommand                                 As New Command
Private mBuyAgainCommand                            As New Command
Private mCloseoutCommand                            As New Command
Private mContractCommand                            As New Command
Private mEndBracketCommand                          As New Command
Private mEndOrdersCommand                           As New Command
Private mEntryCommand                               As New Command
Private mExitCommand                                As New Command
Private mGroupCommand                               As New Command
Private mHelpCommand                                As New Command
Private mHelp1Command                               As New Command
Private mListCommand                                As New Command
Private mPurgeCommand                               As New Command
Private mQuitCommand                                As New Command
Private mQuoteCommand                               As New Command
Private mResetCommand                               As New Command
Private mRolloverCommand                            As New Command
Private mSellCommand                                As New Command
Private mSellAgainCommand                           As New Command
Private mSetBalanceComand                           As New Command
Private mSetFundsCommand                            As New Command
Private mSetGroupFundsCommand                       As New Command
Private mSetGroupRolloverCommand                    As New Command
Private mSetRolloverCommand                         As New Command
Private mShowBalanceComand                          As New Command
Private mStageOrdersCommand                         As New Command
Private mStopLossCommand                            As New Command
Private mTargetCommand                              As New Command

Private mCommands()                                 As Command
Private mCommandsNextIndex                          As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mCommands(31) As Command

addCommand mBatchOrdersCommand, BatchOrdersCommandName
addCommand mBracketCommand, BracketCommandName
addCommand mBuyCommand, BuyCommandName
addCommand mBuyAgainCommand, BuyAgainCommandName
addCommand mCloseoutCommand, CloseoutCommandName
addCommand mContractCommand, ContractCommandName
addCommand mEndBracketCommand, EndBracketCommandName
addCommand mEndOrdersCommand, EndOrdersCommandName
addCommand mEntryCommand, EntryCommandName
addCommand mExitCommand, ExitCommandName
addCommand mGroupCommand, GroupCommandName
addCommand mHelpCommand, HelpCommandName
addCommand mHelp1Command, Help1CommandName
addCommand mListCommand, ListCommandName
addCommand mPurgeCommand, PurgeCommandName
addCommand mQuitCommand, QuitCommandName
addCommand mQuoteCommand, QuoteCommandName
addCommand mResetCommand, ResetCommandName
addCommand mRolloverCommand, RolloverCommandName
addCommand mSellCommand, SellCommandName
addCommand mSellAgainCommand, SellAgainCommandName
addCommand mSetBalanceComand, SetBalanceCommandName
addCommand mSetFundsCommand, SetFundsCommandName
addCommand mSetGroupFundsCommand, SetGroupFundsCommandName
addCommand mSetGroupRolloverCommand, SetGroupRolloverCommandName
addCommand mSetRolloverCommand, SetRolloverCommandName
addCommand mShowBalanceComand, ShowBalanceCommandName
addCommand mStageOrdersCommand, StageOrdersCommandName
addCommand mStopLossCommand, StopLossCommandName
addCommand mTargetCommand, TargetCommandName

ReDim Preserve mCommands(mCommandsNextIndex - 1) As Command
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BatchOrdersCommand() As Command
Set BatchOrdersCommand = mBatchOrdersCommand
End Property

Friend Property Get BracketCommand() As Command
Set BracketCommand = mBracketCommand
End Property

Friend Property Get BuyCommand() As Command
Set BuyCommand = mBuyCommand
End Property

Friend Property Get BuyAgainCommand() As Command
Set BuyAgainCommand = mBuyAgainCommand
End Property

Friend Property Get CloseoutCommand() As Command
Set CloseoutCommand = mCloseoutCommand
End Property

Friend Property Get ContractCommand() As Command
Set ContractCommand = mContractCommand
End Property

Friend Property Get EndBracketCommand() As Command
Set EndBracketCommand = mEndBracketCommand
End Property

Friend Property Get EndOrdersCommand() As Command
Set EndOrdersCommand = mEndOrdersCommand
End Property

Friend Property Get EntryCommand() As Command
Set EntryCommand = mEntryCommand
End Property

Friend Property Get ExitCommand() As Command
Set ExitCommand = mExitCommand
End Property

Friend Property Get GroupCommand() As Command
Set GroupCommand = mGroupCommand
End Property

Friend Property Get HelpCommand() As Command
Set HelpCommand = mHelpCommand
End Property

Friend Property Get Help1Command() As Command
Set Help1Command = mHelp1Command
End Property

Friend Property Get ListCommand() As Command
Set ListCommand = mListCommand
End Property

Friend Property Get PurgeCommand() As Command
Set PurgeCommand = mPurgeCommand
End Property

Friend Property Get QuitCommand() As Command
Set QuitCommand = mQuitCommand
End Property

Friend Property Get QuoteCommand() As Command
Set QuoteCommand = mQuoteCommand
End Property

Friend Property Get ResetCommand() As Command
Set ResetCommand = mResetCommand
End Property

Friend Property Get RolloverCommand() As Command
Set RolloverCommand = mRolloverCommand
End Property

Friend Property Get SellCommand() As Command
Set SellCommand = mSellCommand
End Property

Friend Property Get SellAgainCommand() As Command
Set SellAgainCommand = mSellAgainCommand
End Property

Friend Property Get SetBalanceCommand() As Command
Set SetBalanceCommand = mSetBalanceComand
End Property

Friend Property Get SetFundsCommand() As Command
Set SetFundsCommand = mSetFundsCommand
End Property

Friend Property Get SetGroupFundsCommand() As Command
Set SetGroupFundsCommand = mSetGroupFundsCommand
End Property

Friend Property Get SetGroupRolloverCommand() As Command
Set SetGroupRolloverCommand = mSetGroupRolloverCommand
End Property

Friend Property Get SetRolloverCommand() As Command
Set SetRolloverCommand = mSetRolloverCommand
End Property

Friend Property Get ShowBalanceCommand() As Command
Set ShowBalanceCommand = mShowBalanceComand
End Property

Friend Property Get StageOrdersCommand() As Command
Set StageOrdersCommand = mStageOrdersCommand
End Property

Friend Property Get StopLossCommand() As Command
Set StopLossCommand = mStopLossCommand
End Property

Friend Property Get TargetCommand() As Command
Set TargetCommand = mTargetCommand
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function ParseCommand(ByVal pCommandName As String) As Command
Const ProcName As String = "ParseCommand"
On Error GoTo Err

Dim i As Long
For i = 0 To UBound(mCommands)
    If mCommands(i) Is Nothing Then Exit Function
    If mCommands(i).Parse(pCommandName) Then
        Set ParseCommand = mCommands(i)
        Exit Function
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addCommand(ByVal pCommand As Command, ByVal pCommandName As String)
Const ProcName As String = "addCommand"
On Error GoTo Err

pCommand.Initialise pCommandName

If mCommandsNextIndex > UBound(mCommands) Then
    ReDim Preserve mCommands(2 * (UBound(mCommands) + 1) - 1) As Command
End If

Set mCommands(mCommandsNextIndex) = pCommand
mCommandsNextIndex = mCommandsNextIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



