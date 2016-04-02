VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IStrategyHostView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Public Enum ErrorSeverities
    ErrorSeverityNone
    ErrorSeverityInformation
    ErrorSeverityWarning
    ErrorSeverityCritical
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IStrategyHostView"

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

Public Property Get Strategy() As IStrategy
End Property

Public Property Get Parameters() As Parameters
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStudyToChart( _
                ByVal pChartIndex As Long, _
                ByVal pStudy As IStudy, _
                ByVal pStudyValueNames As EnumerableCollection)
End Sub

Public Function AddTimeframe(ByVal pTimeframe As Timeframe) As Long
End Function

Public Sub ClearPriceAndProfitFields()
End Sub

' no pIndex => all charts
Public Sub DisablePriceDrawing(Optional ByVal pTimeframeIndex As Long)
End Sub

Public Sub DisableProfitDrawing()
End Sub

Public Sub DisableStart()
End Sub

Public Sub DisableTradeDrawing()
End Sub

' no pIndex => all charts
Public Sub EnablePriceDrawing(Optional ByVal pTimeframeIndex As Long)
End Sub

Public Sub EnableProfitDrawing()
End Sub

Public Sub EnableStart()
End Sub

Public Sub EnableTradeDrawing()
End Sub

Public Sub NotifyError(ByVal pTitle As String, ByVal pMessage As String, ByVal pSeverity As ErrorSeverities)
End Sub

Public Sub NotifyBracketOrderProfile(ByRef Value As BracketOrderProfile)
End Sub

Public Sub NotifyEventsPerSecond(ByVal Value As Long)
End Sub

Public Sub NotifyEventsPlayed(ByVal Value As Long)
End Sub

Public Sub NotifyMicrosecsPerEvent(ByVal Value As Long)
End Sub

Public Sub NotifyNewTradeBar(ByVal pBarNumber As Long, ByVal pTimestamp As Date)
End Sub

Public Sub NotifyPosition(ByVal Value As Long)
End Sub

Public Sub NotifyReplayProgress( _
                ByVal pTickfileTimestamp As Date, _
                ByVal pEventsPlayed As Long, _
                ByVal pPercentComplete As Single)
End Sub

Public Sub NotifySessionDrawdown(ByVal Value As Double)
End Sub

Public Sub NotifySessionMaxProfit(ByVal Value As Double)
End Sub

Public Sub NotifySessionProfit(ByVal Value As Double, ByVal pTimestamp As Date)
End Sub

Public Sub NotifyTick(ev As GenericTickEventData)
End Sub

Public Sub NotifyTickerCreated()
End Sub

Public Sub NotifyTickfileCompleted(ByVal pIndex As Long)
End Sub

Public Sub ResetProfitChart()
End Sub

Public Sub ResetTradeChart()
End Sub

Public Sub ShowTradeLine( _
                ByVal pStartTime As Date, _
                ByVal pEndTime As Date, _
                ByVal pEntryPrice As Double, _
                ByVal pExitPrice As Double, _
                ByVal pProfit As Double)
End Sub

Public Sub UpdateLastChartBars()
End Sub

Public Sub WriteLogText(ByVal pMessage As String)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




