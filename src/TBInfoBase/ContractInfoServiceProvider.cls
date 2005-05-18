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

Private Const MaxQueuedRequests As Long = 100

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

Private mRequests(MaxQueuedRequests - 1) As IContracts
Private mFirstRequestID As Long
Private mNextRequestID As Long
Private mRequestsOutstanding As Long

Private mName As String

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mInstrumentFactory As TradingDO.cInstrumentFactory


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = App.ProductName
Set mTimer = New IntervalTimer
mTimer.RepeatNotifications = False
mTimer.TimerIntervalMillisecs = 1
End Sub

'================================================================================
' ICommonServiceProvider Interface Members
'================================================================================

Private Property Let ICommonServiceProvider_CommonServiceConsumer( _
                            ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
mCommonServiceConsumer.RegisterServiceProvider ServiceProviderTypes.ContractInfo, _
                                            Me
End Property

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
ICommonServiceProvider_Name = Name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
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

If mNextRequestID = mFirstRequestID And mRequestsOutstanding > 0 Then
    pContracts.ServiceProviderCantHandle "Too many contract details requests queued", _
                                mName
End If

Set mRequests(mNextRequestID) = pContracts

If mRequestsOutstanding = 0 Then
    ' no requests currently queued so start the timer
    mTimer.StartTimer
End If

mNextRequestID = (mNextRequestID + 1) Mod MaxQueuedRequests
mRequestsOutstanding = mRequestsOutstanding + 1
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
Dim instrumentSummaries As TradingDO.cInstrumentSummaries
Dim instrumentSummary As TradingDO.cInstrumentSummary
Dim instrument As TradingDO.cInstrument
Dim Contract As IContract
Dim OrderTypes() As TradeBuildSP.OrderTypes
Dim validExchanges() As String

Set mInstrumentFactory = New cInstrumentFactory

With mRequests(mFirstRequestID).ContractSpecifier
    Set instrumentSummaries = mInstrumentFactory.loadEx( _
                                    .Symbol, _
                                    secTypeToString(.SecType), _
                                    .Expiry, _
                                    .Exchange, _
                                    .CurrencyCode)
End With

For Each instrumentSummary In instrumentSummaries
    Set instrument = mInstrumentFactory.loadByID(instrumentSummary.id)
    Set Contract = mCommonServiceConsumer.NewContract
    Contract.Specifier = mCommonServiceConsumer.NewContractSpecifier
    
    With Contract.Specifier
        .Symbol = instrument.Symbol
        .SecType = secTypeFromString(instrument.category)
        .Expiry = instrument.Month
        .Exchange = instrument.Exchange
        .CurrencyCode = instrument.CurrencyCode
        .LocalSymbol = instrument.shortName
    End With
    
    With Contract
        .Description = instrument.Name
        .MarketName = instrument.Symbol
        .MinimumTick = instrument.TickSize
        .Multiplier = instrument.TickValue / instrument.TickSize
        .SessionEndTime = instrument.SessionEndTime
        .SessionStartTime = instrument.SessionStartTime
        
        ReDim OrderTypes(3) As TradeBuildSP.OrderTypes
        OrderTypes(0) = TradeBuildSP.OrderTypes.OrderTypeMarket
        OrderTypes(1) = TradeBuildSP.OrderTypes.OrderTypeLimit
        OrderTypes(2) = TradeBuildSP.OrderTypes.OrderTypeStop
        OrderTypes(3) = TradeBuildSP.OrderTypes.OrderTypeStopLimit
        .OrderTypes = OrderTypes
        
        ReDim validExchanges(0) As String
        validExchanges(0) = instrument.Exchange
        .validExchanges = validExchanges
    End With
    
    mRequests(mFirstRequestID).AddContract Contract
Next

mContractInfoServiceConsumer.LoadComplete mHandle, mRequests(mFirstRequestID)

mRequestsOutstanding = mRequestsOutstanding - 1
mFirstRequestID = (mFirstRequestID + 1) Mod MaxQueuedRequests
If mRequestsOutstanding <> 0 Then mTimer.StartTimer
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


