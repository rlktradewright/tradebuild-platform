VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickDataEncoder"
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

Private Const ProjectName                   As String = "TickUtils26"
Private Const ModuleName                    As String = "TickDataEncoder"

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

Public Sub encodeAsk( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
End Sub

Public Sub encodeBid( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
End Sub

Public Sub encodeClose( _
                ByVal timestamp As Date, _
                ByVal price As Double)
End Sub

Public Sub encodeHigh( _
                ByVal timestamp As Date, _
                ByVal price As Double)
End Sub

Public Sub encodeLow( _
                ByVal timestamp As Date, _
                ByVal price As Double)
End Sub

Public Sub encodeMarketDepthData( _
                ByVal timestamp As Date, _
                ByVal position As Long, _
                ByVal marketmaker As String, _
                ByVal operation As Long, _
                ByVal side As Long, _
                ByVal price As Double, _
                ByVal size As Long)
End Sub

Public Sub encodeMarketDepthReset( _
                ByVal timestamp As Date)
End Sub

Public Sub encodeOpenInterest( _
                ByVal timestamp As Date, ByVal size As Long)
End Sub

Public Sub encodeTrade( _
                ByVal timestamp As Date, _
                ByVal price As Double, _
                ByVal size As Long)
End Sub

Public Sub encodeVolume( _
                ByVal timestamp As Date, _
                ByVal size As Long)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

