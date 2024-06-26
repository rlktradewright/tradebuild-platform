VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClosePositionsResult"
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

Private Const ModuleName                            As String = "ClosePositionsResult"

'@================================================================================
' Member variables
'@================================================================================

Private mNumberOfLivePositionsClosed                As Long
Private mNumberOfLivePositionsNotClosed             As Long

Private mNumberOfSimulatedPositionsClosed           As Long
Private mNumberOfSimulatedPositionsNotClosed        As Long

Private mSimulatedErrorMessages                     As New Collection
Private mLiveErrorMessages                          As New Collection

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

Public Property Get LiveErrorMessages() As Collection
Set LiveErrorMessages = mLiveErrorMessages
End Property

Friend Property Let NumberOfLivePositionsClosed(ByVal Value As Long)
mNumberOfLivePositionsClosed = Value
End Property

Public Property Get NumberOfLivePositionsClosed() As Long
NumberOfLivePositionsClosed = mNumberOfLivePositionsClosed
End Property

Friend Property Let NumberOfLivePositionsNotClosed(ByVal Value As Long)
mNumberOfLivePositionsNotClosed = Value
End Property

Public Property Get NumberOfLivePositionsNotClosed() As Long
NumberOfLivePositionsNotClosed = mNumberOfLivePositionsNotClosed
End Property

Friend Property Let NumberOfSimulatedPositionsClosed(ByVal Value As Long)
mNumberOfSimulatedPositionsClosed = Value
End Property

Public Property Get NumberOfSimulatedPositionsClosed() As Long
NumberOfSimulatedPositionsClosed = mNumberOfSimulatedPositionsClosed
End Property

Friend Property Let NumberOfSimulatedPositionsNotClosed(ByVal Value As Long)
mNumberOfSimulatedPositionsNotClosed = Value
End Property

Public Property Get NumberOfSimulatedPositionsNotClosed() As Long
NumberOfSimulatedPositionsNotClosed = mNumberOfSimulatedPositionsNotClosed
End Property

Public Property Get SimulatedErrorMessages() As Collection
Set SimulatedErrorMessages = mSimulatedErrorMessages
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddLiveErrorMessage(ByVal pMessage As String)
mLiveErrorMessages.Add pMessage
End Sub

Friend Sub AddSimulatedErrorMessage(ByVal pMessage As String)
mSimulatedErrorMessages.Add pMessage
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




