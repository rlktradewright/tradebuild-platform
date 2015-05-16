VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamBasedProcessor"
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


Private Const ModuleName                    As String = "StreamBasedProcessor"

'@================================================================================
' Member variables
'@================================================================================

Private mSelfRef As Object

Private WithEvents mTimeframe As Timeframe
Attribute mTimeframe.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' mTimeframe Event Handlers
'@================================================================================

Private Sub mTimeframe_BarsLoaded()
Const ProcName As String = "mTimeframe_BarsLoaded"
On Error GoTo Err

gProcessor.Finish
Set gProcessor = Nothing
gCon.WriteLineToConsole "Completed"
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub StartData( _
                ByVal pTB As TradeBuildAPI, _
                ByVal spec As IContractSpecifier, _
                ByVal fromDate As Date, _
                ByVal toDate As Date, _
                ByVal number As Long, _
                ByVal barLength As Long, _
                ByVal barUnits As TimePeriodUnits, _
                ByVal sessionOnly As Boolean)
Const ProcName As String = "startData"
On Error GoTo Err

Set mSelfRef = Me

Dim lticker As Ticker
Set lticker = pTB.Tickers.CreateTicker(FetchContract(spec, pTB.ContractStorePrimary))
gProcessor.Ticker = lticker

Set mTimeframe = lticker.Timeframes.AddHistorical(GetTimePeriod(barLength, barUnits), _
                                        "", _
                                        IIf(number = -1, 0, number), _
                                        fromDate, _
                                        toDate, _
                                        , _
                                        Not sessionOnly)
gProcessor.Timeframe = mTimeframe

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

