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

Private mTicker                                     As Ticker

Private mStrategy                                   As IStrategy
Private mStopStrategyFactory                        As IPosnMgmtStrtgyFactory

Private mParams                                     As Parameters

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

mStrategy.Initialise pContext, mParams, mStopStrategyFactory

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialisationCompleted()
Const ProcName As String = "InitialisationCompleted"
On Error GoTo Err

Dim lTradingContext As New TradingContext
lTradingContext.Initialise Me, mStrategy, mTicker
mStrategy.Start lTradingContext

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Initialise( _
                ByVal pStrategy As IStrategy, _
                ByVal pStopStrategyFactory As IPosnMgmtStrtgyFactory) As Parameters
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStrategy = pStrategy
Set Initialise = mergeParameters(pStrategy.DefaultParameters, pStopStrategyFactory.DefaultParameters)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Start(ByVal pTicker As Ticker, _
                ByVal pParams As Parameters)
Const ProcName As String = "Start"
On Error GoTo Err

Set mTicker = pTicker
Set mParams = pParams

Dim lContext As New InitialisationContext
lContext.Initialise Me, mTicker

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' Helper Functions
'@================================================================================

Private Function mergeParameters(ByVal pParams1 As Parameters, ByVal pParams2 As Parameters) As Parameters
Const ProcName As String = "mergeParameters"
On Error GoTo Err

Const DefaultValue As String = "***$Default$***"

Dim lParams As New Parameters

Dim lParam As Parameter
For Each lParam In pParams1
    lParams.SetParameterValue lParam.Name, lParam.Value
Next

For Each lParam In pParams2
    If lParams.GetParameterValue(lParam.Name, DefaultValue) = DefaultValue Then lParams.SetParameterValue lParam.Name, lParam.Value
Next

Set mergeParameters = lParams

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
