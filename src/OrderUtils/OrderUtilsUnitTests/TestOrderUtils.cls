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

Private mContractStore                              As IContractStore

Private mMarketDataManager                          As IMarketDataManager

Private mTestCompleted                              As Boolean

'@================================================================================
' Setup Functions
'@================================================================================

Public Sub FixtureSetup()
End Sub

Public Sub FixtureTeardown()
End Sub

Public Sub Setup()
Set mFutureWaiter = New FutureWaiter
Set mContractStore = New MockContractStore
Set mMarketDataManager = New MockMarketDataManager

mTestCompleted = False
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

End Sub

'@================================================================================
' Test Methods
'@================================================================================

Public Sub GetOptionContract_0100()
Dim lContractSpec As IContractSpecifier
mFutureWaiter.Add GetOptionContract(lContractSpec, _
                                    OrderActionBuy, _
                                    mContractStore, _
                                    OptionStrikeSelectionModeExpenditure, _
                                    800, _
                                    OptionStrikeSelectionOperatorLE, _
                                    "SMARTUS", _
                                    mMarketDataManager), "0100"

Do
    Wait 50
Loop Until mTestCompleted

End Sub

'@================================================================================
' Helper Functions
'@================================================================================



