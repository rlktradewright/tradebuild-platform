VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOrderSubmissionSrvcConsumer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Get clockRate() As Single

End Property

'================================================================================
' Methods
'================================================================================

Public Sub cancelTickData(ByVal consumer As TradeBuildsp.IStreamingDataConsumer)

End Sub

Public Sub executionDetails( _
                ByVal orderId As String, _
                ByVal pContractSpecifier As TradeBuildsp.IContractSpecifier, _
                ByVal exec As TradeBuild.Execution)
End Sub

Public Sub orderStatus(ByVal orderId As String, _
                ByVal status As TradeBuildsp.OrderStatuses, _
                ByVal filled As Long, _
                ByVal remaining As Long, _
                ByVal avgFillPrice As Double, _
                ByVal lastFillPrice As Double)
End Sub

Public Sub preFill( _
                ByVal orderId As Long, _
                ByRef fillPrice As Double, _
                ByRef fillSize As Long)
End Sub

Public Sub rejectOrder( _
                ByVal orderId As String, _
                ByVal reason As String)
End Sub

Public Sub requestTickData( _
                ByVal consumer As TradeBuildsp.IStreamingDataConsumer, _
                ByVal includeMarketDepth As Boolean)
End Sub

'================================================================================
' Helper Functions
'================================================================================

