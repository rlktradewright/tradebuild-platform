VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSubmissionSrvcProvider"
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

Private mOrderSubmitters As Collection

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mOrderSubmitters = New Collection

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
On Error GoTo err
gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.LiveOrderSubmission
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmissionServiceProvider" & "::" & "ICommonServiceProvider_Link", _
                err.Description

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
mLogLevel = RHS
End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
finishSubmitters
Set mCommonServiceConsumer = Nothing
Exit Sub
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmissionServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal contract As TradeBuildSP.IContract) As TradeBuildSP.IOrderSubmitter
Dim orderSubmitter As orderSubmitter

On Error GoTo err
Set orderSubmitter = New orderSubmitter
orderSubmitter.OrderSubmissionServiceProvider = Me
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
mOrderSubmitters.add orderSubmitter

orderSubmitter.start

Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = orderSubmitter
Exit Function
err:
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmissionServiceProvider" & "::" & "IOrderSubmissionSrvcProvider_CreateOrderSubmitter", _
                err.Description
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

Friend Property Get handle() As Long
handle = mHandle
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

Friend Sub orderSubmitterFinished( _
                ByVal pOrderSubmitter As orderSubmitter)
Dim i As Long
Dim submitter As orderSubmitter
i = 1
For Each submitter In mOrderSubmitters
    If submitter Is pOrderSubmitter Then
        mOrderSubmitters.remove i
        Exit For
    End If
    i = i + 1
Next
End Sub

'===============================================================================
' Helper Functions
'================================================================================

Private Sub finishSubmitters()
Dim submitter As orderSubmitter
For Each submitter In mOrderSubmitters
    submitter.finish
Next
End Sub

Private Sub handleFatalError( _
                ByVal number As Long, _
                ByVal source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError number, source, Description, mHandle

finishSubmitters

Set mCommonServiceConsumer = Nothing
End Sub



