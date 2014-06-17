VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface must be implemented by classes that operate as trading strategies
' within the TradeBuild Strategy Host program.
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

Private Const ModuleName                    As String = "IStrategy"

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

Public Property Get DefaultParameters() As Parameters
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Called to enable the <code>Strategy</code> object to initialise itself, for example
' by creating any necessary timeframes and studies.
'
' @param StrategyContext
'   The <code>StrategyContext</code> object to be used by this <code>Strategy</code>.
' @param pParams
'   The parameters for this <code>Strategy</code>.
'@/
Public Sub Initialise( _
                ByVal pInitialisationContext As InitialisationContext, _
                ByVal pParams As Parameters, _
                ByVal pStopStrategyFactory As IPosnMgmtStrtgyFactory)
End Sub

''
' Called for each tick in the underlying <code>Ticker</code> object.
'
' @remarks
'   The relevant prices and sizes relating to this tick can be obtained using the
'   revelant properties of the underlying <code>Ticker</code> object, such as
'   <code>AskPrice</code>, <code>AskPriceString</code>, <code>AskSize</code> etc.
' @param pType
'   The type of tick that has occurred.
'@/
Public Sub NotifyTick(ByVal pType As TickTypes)

End Sub

''
' Called after historic price data has been loaded,
' but before the first tick is notified.
Public Sub Start(ByVal pTradingContext As TradingContext)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


