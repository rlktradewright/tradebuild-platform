VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractInfoServiceProvider"
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

Implements ICommonServiceProvider
Implements IContractInfoServiceProvider

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const InitialRequestQueueCapacity As Long = 100

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mHandle As Long
Private mContractInfoServiceConsumer As IContractInfoServiceConsumer

Private mRequestQueue As Queue

Private mCurrentRequestedContracts As TradeBuildSP.IContracts

Private mName As String

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private WithEvents mTWSAPI As TWSAPI
Attribute mTWSAPI.VB_VarHelpID = -1
Private mClientID As Long
Private mServer As String
Private mPort As Long
Private mConnectionRetryIntervalSecs As Long

Private mLogLevel As LogLevels
Private mTWSLogLevel As TWSLogLevels

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = ContractInfoSPName
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 1

mClientID = -1
mPort = 7496
mConnectionRetryIntervalSecs = 10
mLogLevel = LogLevelLow
mTWSLogLevel = TWSLogLevelError

Set mRequestQueue = New Queue
mRequestQueue.initialCapacity = InitialRequestQueueCapacity
End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim ServiceProviderDetails  As TradeBuildSP.ServiceProviderDetails
With ServiceProviderDetails
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
ICommonServiceProvider_Details = ServiceProviderDetails
End Property

Private Sub ICommonServiceProvider_Link( _
                            ByVal commonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                            ByVal handle As Long)
Set mCommonServiceConsumer = commonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider handle, _
                                            ServiceProviderTypes.ContractInfo
End Sub

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
If Not mTWSAPI Is Nothing Then gReleaseTWSAPIInstance mTWSAPI
Set mTWSAPI = Nothing
End Sub

'================================================================================
' IContractInfoServiceProvider Interface Members
'================================================================================

Private Property Let IContractInfoServiceProvider_ContractInfoServiceConsumer( _
                            ByVal RHS As TradeBuildSP.IContractInfoServiceConsumer)
Set mContractInfoServiceConsumer = RHS
End Property

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                            ByVal pContracts As TradeBuildSP.IContracts)

If mTWSAPI Is Nothing Then
    If mClientID = -1 Then mClientID = Int(&H7FFFFFFF * Rnd) + 1
    Set mTWSAPI = gGetTWSAPIInstance(mServer, mPort, mClientID, mConnectionRetryIntervalSecs, True)
    mTWSAPI.commonServiceConsumer = mCommonServiceConsumer
    mTWSAPI.logLevel = mLogLevel
    mTWSAPI.TWSLogLevel = mTWSLogLevel
End If

mRequestQueue.addToEnd pContracts

If mRequestQueue.size = 1 And mTWSAPI.connectionState = ConnConnected Then
    ' no requests previously queued so start the timer
    mTimer.StartTimer
End If

End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
                            ByVal pContract As TradeBuildSP.IContract)

End Sub

Private Function IContractInfoServiceProvider_Supports( _
                            ByVal capabilities As Long) As Boolean
Dim supportedCapabilities As Long

supportedCapabilities = ContractInfoServiceProviderCapabilities.Store
IContractInfoServiceProvider_Supports = supportedCapabilities And capabilities

End Function

'================================================================================
' mTimer Event Handlers
'================================================================================

Private Sub mTimer_TimerExpired()
Set mCurrentRequestedContracts = mRequestQueue.removeFromFront
mTWSAPI.RequestContract mCurrentRequestedContracts
End Sub

'================================================================================
' mTWSAPI Event Handlers
'================================================================================

Private Sub mTWSAPI_Connected()
If mRequestQueue.size <> 0 Then mTimer.StartTimer
End Sub

Private Sub mTWSAPI_ConnectFailed(ByVal Description As String, ByVal retrying As Boolean)
mTimer.StopTimer
If Not retrying Then cancelRequests "Can't connect to TWS"
End Sub

Private Sub mTWSAPI_ConnectionToIBClosed()
mTimer.StopTimer
End Sub

Private Sub mTWSAPI_ConnectionToIBRecovered()
If mRequestQueue.size <> 0 Then mTimer.StartTimer
End Sub

Private Sub mTWSAPI_ConnectionToTWSClosed(ByVal reconnecting As Boolean)
mTimer.StopTimer
If Not reconnecting Then cancelRequests "Connection to TWS closed"
End Sub

Private Sub mTWSAPI_ContractSpecifierInvalid( _
                ByVal reason As String, _
                ByVal pContracts As TradeBuildSP.IContracts)
mContractInfoServiceConsumer.Error StandardSPErrorCodes.CIContractSpecifierInvalid, _
                                    reason, _
                                    pContracts
If mRequestQueue.size <> 0 Then mTimer.StartTimer
End Sub

Private Sub mTWSAPI_ContractsLoaded()
mContractInfoServiceConsumer.LoadComplete mHandle, mCurrentRequestedContracts
If mRequestQueue.size <> 0 Then mTimer.StartTimer
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let clientID(ByVal value As Long)
mClientID = value
End Property

Public Property Let connectionRetryIntervalSecs(ByVal value As Long)
mConnectionRetryIntervalSecs = value
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

Public Property Let server(ByVal value As String)
mServer = IIf(value = "", "127.0.0.1", value)
End Property

Public Property Let TWSLogLevel(ByVal value As TWSLogLevels)
Select Case value
Case TWSLogLevelSystem
Case TWSLogLevelError
Case TWSLogLevelWarning
Case TWSLogLevelInformation
Case TWSLogLevelDetail
Case Else
    err.Raise ErrorCodes.IllegalArgumentException, _
                "IBTWSSP.ContractInfoServiceProvider::(let)TWSLogLevel", _
                "Value must be one of the TWSLogLevels enum"
End Select
mTWSLogLevel = value
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub cancelRequests(ByVal reason As String)
Dim lContracts As TradeBuildSP.IContracts
Do While mRequestQueue.size <> 0
    Set lContracts = mRequestQueue.removeFromFront
    lContracts.ServiceProviderCantHandle reason, mName
Loop


End Sub



