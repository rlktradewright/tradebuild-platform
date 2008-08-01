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
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

gCommonServiceConsumer = commonServiceConsumer
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.LiveOrderSubmission

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(ParamNameClientId)
        clientID = gParseClientId(param.value)
    Case UCase$(ParamNameConnectionRetryIntervalSecs)
        connectionRetryIntervalSecs = gParseConnectionRetryInterval(param.value)
    Case UCase$(ParamNameKeepConnection)
        keepConnection = gParseKeepConnection(param.value)
    Case UCase$(ParamNamePort)
        port = gParsePort(param.value)
    Case UCase$(ParamNameProviderKey)
        providerKey = param.value
    Case UCase$(ParamNameServer)
        server = param.value
    Case UCase$(ParamNameTwsLogLevel)
        TWSLogLevel = gParseTwsLogLevel(param.value)
    End Select
Next

End Sub

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = mName
End Property

Private Sub ICommonServiceProvider_Terminate()
On Error GoTo err
finishSubmitters
Set mCommonServiceConsumer = Nothing
Exit Sub
err:
logMessage "IBTWSSP:OrderSubmissionSrvcProvider:ICommonServiceProvider_Terminate (" & err.number & "): " & err.Description, LogLevelSevere
handleFatalError err.number, _
                "IBTWSSP" & "." & "OrderSubmissionServiceProvider" & "::" & "ICommonServiceProvider_Terminate", _
                err.Description
End Sub

'================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal pContract As Contract) As TradeBuildSP.IOrderSubmitter

Dim OrderSubmitter As OrderSubmitter

On Error GoTo err
Set OrderSubmitter = New OrderSubmitter
OrderSubmitter.OrderSubmissionServiceProvider = Me
OrderSubmitter.commonServiceConsumer = mCommonServiceConsumer
OrderSubmitter.serviceConsumer = serviceConsumer
OrderSubmitter.Contract = pContract
OrderSubmitter.server = mServer
OrderSubmitter.port = mPort
OrderSubmitter.clientID = mClientID
OrderSubmitter.connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
OrderSubmitter.keepConnection = mkeepConnection
OrderSubmitter.providerKey = mProviderKey
OrderSubmitter.TWSLogLevel = mTWSLogLevel
mOrderSubmitters.add OrderSubmitter

OrderSubmitter.start

Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = OrderSubmitter
Exit Function
err:
logMessage "IBTWSSP:OrderSubmissionSrvcProvider:IOrderSubmissionSrvcProvider_CreateOrderSubmitter (" & err.number & "): " & err.Description, LogLevelSevere
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

Friend Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Friend Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
End Property

Friend Property Get connectionRetryIntervalSecs() As Long
connectionRetryIntervalSecs = mConnectionRetryIntervalSecs
End Property

Friend Property Get handle() As Long
handle = mHandle
End Property

Friend Property Let keepConnection(ByVal value As Boolean)
mkeepConnection = value
End Property

Friend Property Let name(ByVal value As String)
mName = value
End Property

Friend Property Let port(ByVal value As Long)
mPort = value
End Property

Friend Property Let providerKey(ByVal value As String)
mProviderKey = value
End Property

Friend Property Let server(ByVal value As String)
mServer = value
If mServer = "" Then
    mServer = "127.0.0.1"
End If
End Property

Friend Property Let TWSLogLevel(ByVal value As TWSLogLevels)
Select Case value
Case TWSLogLevelSystem
Case TWSLogLevelError
Case TWSLogLevelWarning
Case TWSLogLevelInformation
Case TWSLogLevelDetail
Case Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "IBTWSSP.RealtimeDataServiceProvider::(let)TWSLogLevel", _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

Friend Sub orderSubmitterFinished( _
                ByVal pOrderSubmitter As OrderSubmitter)
Dim i As Long
Dim submitter As OrderSubmitter
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
Dim submitter As OrderSubmitter
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

Private Sub logMessage( _
                ByVal message As String, _
                Optional ByVal logLevel As LogLevels = LogLevelNormal)
If Not gLogger.isLoggable(logLevel) Then Exit Sub
gLogger.Log logLevel, _
            message
End Sub


