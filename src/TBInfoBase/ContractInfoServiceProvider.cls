VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractInfoSrvcProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
                           Option Explicit

'@===============================================================================
' Description
'@===============================================================================
'
'
'@===============================================================================
' Amendment history
'@===============================================================================
'
'
'
'

'@===============================================================================
' Interfaces
'@===============================================================================

Implements ICommonServiceProvider
Implements IContractInfoServiceProvider

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName As String = "TBInfoBase26"
Private Const ModuleName As String = "ContractInfoSrvcProvider"

Private Const MaxQueuedRequests As Long = 100

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private WithEvents mTradingDB       As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString           As String

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mHandle As Long
Private mContractInfoServiceConsumer As IContractInfoServiceConsumer

Private mRequests(MaxQueuedRequests - 1) As ContractsBuilder
Private mFirstRequestID As Long
Private mNextRequestID As Long
Private mRequestsOutstanding As Long

Private mName As String

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mInstrumentFactory As InstrumentFactory

Private mDatabaseType As DatabaseTypes
Private mDatabaseName As String
Private mServer As String
Private mUserName As String
Private mPassword As String

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mName = ContractInfoSPName
Set mTimer = CreateIntervalTimer(1)

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""
End Sub

'@===============================================================================
' ICommonServiceProvider Interface Members
'@===============================================================================

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
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal handle As Long, _
                ByVal params As Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider handle, _
                                            ServiceProviderTypes.ContractInfo

For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameConnectionString)
        connectionString = param.value
    Case UCase$(ParamNameDatabaseType)
        databaseType = DatabaseTypeFromString(param.value)
    Case UCase$(ParamNameDatabaseName)
        databaseName = param.value
    Case UCase$(ParamNameServer)
        server = param.value
    Case UCase$(ParamNameUserName)
        username = param.value
    Case UCase$(ParamNamePassword)
        password = param.value
    End Select
Next

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)

End Property

Private Property Let ICommonServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get ICommonServiceProvider_Name() As String
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'@===============================================================================
' IContractInfoServiceProvider Interface Members
'@===============================================================================

Private Property Let IContractInfoServiceProvider_ContractInfoServiceConsumer( _
                            ByVal RHS As TradeBuildSP.IContractInfoServiceConsumer)
Set mContractInfoServiceConsumer = RHS
End Property

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                            ByVal pContractsBuilder As ContractsBuilder)

If mNextRequestID = mFirstRequestID And mRequestsOutstanding > 0 Then
    pContractsBuilder.unableToLoad "Too many contract details requests queued", _
                                    mName
End If

Set mRequests(mNextRequestID) = pContractsBuilder

If mRequestsOutstanding = 0 Then
    ' no requests currently queued so start the timer
    mTimer.StartTimer
End If

mNextRequestID = (mNextRequestID + 1) Mod MaxQueuedRequests
mRequestsOutstanding = mRequestsOutstanding + 1
End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
                            ByVal pContract As Contract)

End Sub

Private Function IContractInfoServiceProvider_Supports( _
                            ByVal capabilities As Long) As Boolean
Dim supportedCapabilities As Long

supportedCapabilities = ContractInfoServiceProviderCapabilities.Store
IContractInfoServiceProvider_Supports = supportedCapabilities And capabilities

End Function

'@===============================================================================
' mTimer Event Handlers
'@===============================================================================

Private Sub mTimer_TimerExpired()
Dim instrumentSummaries As DataObjectSummaries
Dim instrumentSummary As DataObjectSummary
Dim instrument As instrument
Dim lContract As Contract
Dim fieldNames() As String

Dim failpoint As Long
On Error GoTo Err

If mTradingDB Is Nothing Then
    If mConnectionString <> "" Then
        Set mTradingDB = CreateTradingDB(mConnectionString)
    Else
        Set mTradingDB = CreateTradingDB(GenerateConnectionString(mDatabaseType, _
                                                                mServer, _
                                                                mDatabaseName, _
                                                                mUserName, _
                                                                mPassword))
    End If
End If

Set mInstrumentFactory = mTradingDB.InstrumentFactory

Set instrumentSummaries = mInstrumentFactory.loadSummariesBySpec( _
                                        fieldNames, _
                                        mRequests(mFirstRequestID).Contracts.ContractSpecifier)

For Each instrumentSummary In instrumentSummaries
    Set instrument = mInstrumentFactory.loadByID(instrumentSummary.id)
    
    Set lContract = ContractFromInstrument(instrument)
    
    mRequests(mFirstRequestID).AddContract lContract
Next

mContractInfoServiceConsumer.LoadComplete mHandle, mRequests(mFirstRequestID)

mRequestsOutstanding = mRequestsOutstanding - 1
mFirstRequestID = (mFirstRequestID + 1) Mod MaxQueuedRequests
If mRequestsOutstanding <> 0 Then mTimer.StartTimer

Exit Sub

Err:

mCommonServiceConsumer.FatalServiceProviderError Err.Number, _
                ProjectName & "." & ModuleName & ":" & "mTimer_TimerExpired" & "." & failpoint & _
                IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
                Err.Description, _
                mHandle
End Sub

'@===============================================================================
' Properties
'@===============================================================================

Public Property Let connectionString( _
                ByVal value As String)
mConnectionString = value
End Property

Public Property Let databaseName( _
                ByVal value As String)
mDatabaseName = value
End Property

Public Property Let databaseType( _
                ByVal value As DatabaseTypes)
mDatabaseType = value
End Property

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let password( _
                ByVal value As String)
mPassword = value
End Property
                
Public Property Let server( _
                ByVal value As String)
mServer = value
End Property
                
Public Property Let username( _
                ByVal value As String)
mUserName = value
End Property

'@===============================================================================
' Methods
'@===============================================================================

'@===============================================================================
' Helper Functions
'@===============================================================================


