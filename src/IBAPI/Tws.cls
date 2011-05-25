VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tws"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "Tws"

'@================================================================================
' Member variables
'@================================================================================

Private mTwsApis                                    As New EnumerableCollection

Private mServer                                     As String
Private mPort                                       As String

Private mHistDataConsumer                           As IHistDataConsumer
Private mContractDetailsConsumer                    As IContractDetailsConsumer
Private mOrderInfoConsumer                          As IOrderInfoConsumer
Private mMarketDataConsumer                         As IMarketDataConsumer
Private mMarketDepthConsumer                        As IMarketDepthConsumer

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

Friend Property Let ContractDetailsConsumer(ByVal Value As IContractDetailsConsumer)
Set mContractDetailsConsumer = Value
End Property

Friend Property Get ContractDetailsConsumer() As IContractDetailsConsumer
Set ContractDetailsConsumer = mContractDetailsConsumer
End Property

Friend Property Let HistDataConsumer(ByVal Value As IHistDataConsumer)
Set mHistDataConsumer = Value
End Property

Friend Property Get HistDataConsumer() As IHistDataConsumer
Set HistDataConsumer = mHistDataConsumer
End Property

Friend Property Let MarketDataConsumer(ByVal Value As IMarketDataConsumer)
Set mMarketDataConsumer = Value
End Property

Friend Property Get MarketDataConsumer() As IMarketDataConsumer
Set MarketDataConsumer = mMarketDataConsumer
End Property

Friend Property Let MarketDepthConsumer(ByVal Value As IMarketDepthConsumer)
Set mMarketDepthConsumer = Value
End Property

Friend Property Get MarketDepthConsumer() As IMarketDepthConsumer
Set MarketDepthConsumer = mMarketDepthConsumer
End Property

Friend Property Let OrderInfoConsumer(ByVal Value As IOrderInfoConsumer)
Set mOrderInfoConsumer = Value
End Property

Friend Property Get OrderInfoConsumer() As IOrderInfoConsumer
Set OrderInfoConsumer = mOrderInfoConsumer
End Property

Friend Property Get Port() As String
Port = mPort
End Property

Friend Property Get Server() As String
Server = mServer
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pServer As String, _
                ByVal pPort As Long)
mServer = pServer
mPort = pPort
End Sub

Friend Function GetAPI(ByVal pClientId As Long) As TwsAPI
Const ProcName As String = "GetAPI"
On Error GoTo Err

Dim lKey As String
lKey = CStr(pClientId)

If Not mTwsApis.Contains(lKey) Then
    Set GetAPI = New TwsAPI
    mTwsApis.Add GetAPI, lKey
    
    GetAPI.Initialise Me, pClientId
    GetAPI.Connect
Else
    Set GetAPI = mTwsApis(lKey)
End If

GetAPI.IncrementUsageCount

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Friend Sub ReleaseAPI( _
                ByVal pClient As TwsAPI, _
                ByVal pForceDisconnect As Boolean)
Const ProcName As String = "ReleaseAPI"
On Error GoTo Err

pClient.DecrementUsageCount
If pClient.UsageCount = 0 Or pForceDisconnect Then If pClient.ConnectionState <> ConnNotConnected Then pClient.Disconnect "release", pForceDisconnect

mTwsApis.Remove CStr(pClient.ClientID)
If mTwsApis.Count = 0 Then gReleaseTws Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




