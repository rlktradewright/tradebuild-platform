VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IMarketDataConsumer"
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "IMarketDataConsumer"

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

'@================================================================================
' Methods
'@================================================================================

Public Sub EndTickSnapshot(ByVal pReqId As Long)
End Sub

Public Sub NotifyError( _
                ByVal pTickerId As Long, _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
End Sub

Public Sub NotifyTickEFP( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pBasisPoints As Double, _
                ByVal pFormattedBasisPoints As String, _
                ByVal pTotalDividends As Double, _
                ByVal pHoldDays As Long, _
                ByVal pFutureExpiry As String, _
                ByVal pDividendImpact As Double, _
                ByVal pDividendsToExpiry As Double)
End Sub

Public Sub NotifyTickGeneric( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pValue As Double)
End Sub


Public Sub NotifyTickOptionComputation( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pImpliedVol As Double, _
                ByVal pDelta As Double, _
                ByVal pOptPrice As Double, _
                ByVal pPvDividend As Double, _
                ByVal pGamma As Double, _
                ByVal pVega As Double, _
                ByVal pTheta As Double, _
                ByVal pUndPrice As Double)
End Sub

Public Sub NotifyTickPrice( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pPrice As Double, _
                ByVal pSize As Long, _
                ByRef pAttributes As TwsTickAttributes)
End Sub

Public Sub NotifyTickRequestParams( _
                ByVal pTickerId As Long, _
                ByVal pMinTick As Double, _
                ByVal pBboExchange As String, _
                ByVal pSnapshotPermissions As Long)

End Sub

Public Sub NotifyTickSize( _
                ByVal pTickerId As Long, _
                ByVal pTickType As Long, _
                ByVal pSize As Long)
End Sub

Public Sub NotifyTickString( _
                ByVal pTickerId As Long, _
                ByVal pTickType As TwsTickTypes, _
                ByVal pValue As String)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




