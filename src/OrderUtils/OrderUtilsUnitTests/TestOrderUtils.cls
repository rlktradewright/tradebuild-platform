VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestOrderUtils"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

Private mContractStore                              As IContractStore

Private mMarketDataManager                          As IMarketDataManager

Private mTestCompleted                              As Boolean

Private mContract                                   As IContract

Private mErrorMessage                               As String

Const TestId0100                                    As String = "0100"
Const TestId0200                                    As String = "0200"
Const TestId1000                                    As String = "1000"
Const TestId1500                                    As String = "1500"

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
' nothing to do
End Sub

Public Sub FixtureTeardown()
' nothing to do
End Sub

Public Sub Setup()
Set mFutureWaiter = New FutureWaiter
Set mContractStore = New MockContractStore
Set mMarketDataManager = New MockMarketDataManager

mErrorMessage = ""
mTestCompleted = False

Set mContract = Nothing
End Sub

Public Sub TearDown()
LogMessage "Test Completed -----------------------------------------------"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If ev.Future.IsPending Then Exit Sub

If ev.Future.IsCancelled Then
ElseIf ev.Future.IsFaulted Then
    mTestCompleted = True
Else
    Select Case ev.ContinuationData
    Case TestId0100, _
            TestId0200, _
            TestId1000, _
            TestId1500
        Set mContract = ev.Future.Value
        mTestCompleted = True
    End Select
End If
    
End Sub

'@================================================================================
' Test Methods
'@================================================================================

Public Sub GetOptionContract_0100()
Dim Expiry As String
' contract /symbol:msft /sectype:opt /expiry:0 /exchange:smart /right:c /strike:800$

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(Symbol:="MSFT", _
                                        localSymbol:="MSFT", _
                                        SecType:=SecTypeOption, _
                                        Expiry:="0", _
                                        Exchange:="smart", _
                                        Right:=OptCall, _
                                        Strike:="0")
mFutureWaiter.Add GetOptionContract(lContractSpec, _
                                    OrderActionBuy, _
                                    mContractStore, _
                                    OptionStrikeSelectionModeExpenditure, _
                                    800, _
                                    OptionStrikeSelectionOperatorLE, _
                                    "SMART", _
                                    mMarketDataManager, _
                                    pReferenceDate:=CDate("2020/07/29")), TestId0100

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue mErrorMessage = ""
Assert.IsFalse mContract Is Nothing
If Not mContract Is Nothing Then
    Assert.AreEqual mContract.Specifier.localSymbol, "MSFT  200731C00195000"
End If
End Sub

Public Sub GetOptionContract_0200()

' contract /symbol:msft /sectype:opt /expiry:0 /exchange:smart /right:c /strike:>=800$

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(Symbol:="MSFT", _
                                        localSymbol:="MSFT", _
                                        SecType:=SecTypeOption, _
                                        Expiry:="0", _
                                        Exchange:="smart", _
                                        Right:=OptCall, _
                                        Strike:="0")
mFutureWaiter.Add GetOptionContract(lContractSpec, _
                                    OrderActionBuy, _
                                    mContractStore, _
                                    OptionStrikeSelectionModeExpenditure, _
                                    800, _
                                    OptionStrikeSelectionOperatorGE, _
                                    "SMART", _
                                    mMarketDataManager, _
                                    pReferenceDate:=CDate("2020/07/29")), TestId0100

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue mErrorMessage = ""
Assert.IsFalse mContract Is Nothing
If Not mContract Is Nothing Then
    Assert.AreEqual mContract.Specifier.localSymbol, "MSFT  200731C00192500"
End If
End Sub

Public Sub GetOptionContract_1000()

' contract /symbol:msft /sectype:opt /expiry:0 /exchange:smart /right:c /strike:>=800$

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(Symbol:="MSFT", _
                                        localSymbol:="MSFT", _
                                        SecType:=SecTypeOption, _
                                        Expiry:="0", _
                                        Exchange:="smart", _
                                        Right:=OptCall, _
                                        Strike:="0")
mFutureWaiter.Add GetOptionContract(lContractSpec, _
                                    OrderActionBuy, _
                                    mContractStore, _
                                    OptionStrikeSelectionModeDelta, _
                                    83, _
                                    OptionStrikeSelectionOperatorLE, _
                                    "SMART", _
                                    mMarketDataManager, _
                                    pReferenceDate:=CDate("2020/07/29")), TestId1000

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue mErrorMessage = ""
Assert.IsFalse mContract Is Nothing
If Not mContract Is Nothing Then
    Assert.AreEqual mContract.Specifier.localSymbol, "MSFT  200731C00195000"
End If
End Sub

Public Sub GetOptionContract_1500()

' contract /symbol:msft /sectype:opt /expiry:0 /exchange:smart /right:c /strike:>=800$

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier(Symbol:="MSFT", _
                                        localSymbol:="MSFT", _
                                        SecType:=SecTypeOption, _
                                        Expiry:="0", _
                                        Exchange:="smart", _
                                        Right:=OptCall, _
                                        Strike:="0")
mFutureWaiter.Add GetOptionContract(lContractSpec, _
                                    OrderActionBuy, _
                                    mContractStore, _
                                    OptionStrikeSelectionModeDelta, _
                                    35, _
                                    OptionStrikeSelectionOperatorLE, _
                                    "SMART", _
                                    mMarketDataManager, _
                                    pReferenceDate:=CDate("2020/07/29")), TestId1500

Do
    Wait 50
Loop Until mTestCompleted

Assert.IsTrue mErrorMessage = ""
Assert.IsFalse mContract Is Nothing
If Not mContract Is Nothing Then
    Assert.AreEqual mContract.Specifier.localSymbol, "MSFT  200731C00207500"
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



