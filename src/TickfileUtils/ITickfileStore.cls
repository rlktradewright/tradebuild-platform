VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ITickfileStore"
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

Private Const ModuleName                            As String = "ITickfileStore"

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

Public Property Get SupportedFormats() As TickfileFormatSpecifier()
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function CreateTickfileReader( _
                ByVal pDataConsumer As IGenericTickListener, _
                ByVal pReplayMonitor As ITickfileReplayMonitor, _
                ByVal pStreamId As Long) As ITickfileReader
End Function

Public Function CreateTickfileWriter( _
                ByVal pOutputMonitor As ITickfileOutputMonitor, _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pFormatIdentifier As String = "", _
                Optional ByVal pLocation As String = "") As ITickfileWriter
End Function

Public Sub Finish()
End Sub

Public Function Supports( _
                ByVal pCapabilities As TickfileStoreCapabilities, _
                Optional ByVal pFormatIdentifier As String) As Boolean
End Function

'@================================================================================
' Helper Functions
'@================================================================================




