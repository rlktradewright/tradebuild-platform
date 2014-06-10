VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HostController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides the means for a <code>Strategy</code> object to interact with
' the Strategy Host program running it.
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

Private Const ModuleName                            As String = "HostController"

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                                 As IMarketDataSource
Private mPositionManager                            As PositionManager
Private mPositionManagerSimulated                   As PositionManager

Private mStrategy                                   As IStrategy
Private mStrategyParams                             As Parameters

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

Friend Sub InitialisationContextIsReady(ByVal pContext As InitialisationContext)
Const ProcName As String = "InitialisationContextIsReadyContextIsReady"
On Error GoTo Err

mStrategy.Initialise pContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialisationCompleted()
Const ProcName As String = "InitialisationCompleted"
On Error GoTo Err

Dim lTradingContext As New TradingContext
lTradingContext.Initialise Me, mStrategy, mDataSource, mPositionManager, mPositionManagerSimulated
mStrategy.Start lTradingContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Initialise( _
                ByVal pStrategy As IStrategy, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pPositionManager As PositionManager, _
                ByVal pPositionManagerSimulated As PositionManager, _
                ByVal pTimeframes As Timeframes, _
                ByVal pStudyManager As StudyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategy = pStrategy
Set mStrategyParams = pStrategy.Parameters

Set mDataSource = pDataSource
Set mPositionManager = pPositionManager
Set mPositionManagerSimulated = pPositionManagerSimulated

Dim lContext As New InitialisationContext
lContext.Initialise Me, pDataSource, pTimeframes, pStudyManager

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


