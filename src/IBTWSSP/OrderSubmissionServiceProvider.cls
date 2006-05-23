VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSubmissionServiceProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Implements TradeBuildSP.ICommonServiceProvider
Implements TradeBuildSP.IOrderSubmissionSrvcProvider

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

Private mServer As String
Private mPort As Long
Private mClientID As Long
Private mConnectionRetryIntervalSecs As Long
Private mkeepConnection As Boolean
Private mProviderKey As String

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mTWSLogLevel As TWSLogLevels
Private mLogLevel As LogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = OrderSubmissionSPName

mClientID = -1
mServer = "127.0.0.1"
mPort = 7496
mConnectionRetryIntervalSecs = 10

mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError
End Sub

Private Sub Class_Terminate()
Debug.Print "IBTWSSP.OrderSubmissionServiceProvider terminated"
End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
With details
    .Comments = App.Comments
    .EXEName = App.EXEName
    .FileDescription = App.FileDescription
    .LegalCopyright = App.LegalCopyright
    .LegalTrademarks = App.LegalTrademarks
    .Path = App.Path
    .ProductName = App.ProductName
    .Vendor = App.CompanyName
    .VersionMajor = App.Major
    .VersionMinor = App.Minor
    .VersionRevision = App.Revision
End With
ICommonServiceProvider_Details = details
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.LiveOrderSubmission

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal contract As TradeBuildSP.IContract) As TradeBuildSP.IOrderSubmitter
Dim orderSubmitter As orderSubmitter

Set orderSubmitter = New orderSubmitter
orderSubmitter.commonServiceConsumer = mCommonServiceConsumer
orderSubmitter.serviceConsumer = serviceConsumer
orderSubmitter.contract = contract
orderSubmitter.server = mServer
orderSubmitter.port = mPort
orderSubmitter.clientID = mClientID
orderSubmitter.connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
orderSubmitter.keepConnection = mkeepConnection
orderSubmitter.providerKey = mProviderKey
orderSubmitter.logLevel = mLogLevel
orderSubmitter.TWSLogLevel = mTWSLogLevel

orderSubmitter.start

Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = orderSubmitter
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Public Property Get connectionRetryIntervalSecs() As Long
connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Public Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Public Property Let logLevel(ByVal value As LogLevels)
mLogLevel = value
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Let port(ByVal value As Long)
mPort = value
End Property

Public Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Public Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Public Property Let TWSLogLevel(ByVal value As TWSLogLevels)
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================



