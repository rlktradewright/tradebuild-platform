VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderSimulatorSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mSimulator As TradeBuild.AdvancedOrderSimulator

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
mName = "TradeBuild Default Order Simulator"

Set mSimulator = New TradeBuild.AdvancedOrderSimulator
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderSimulatorSP terminated"
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
mSimulator.commonServiceConsumer = mCommonServiceConsumer
mHandle = handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.ExchangeSimulator

End Sub

Private Property Let ICommonServiceProvider_LogLevel( _
                ByVal RHS As TradeBuildSP.LogLevels)
End Property

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)

End Property

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = name
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
End Sub

'================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal Contract As TradeBuildSP.IContract) As TradeBuildSP.IOrderSubmitter
Dim orderSubmitter As TradeBuild.SimulatedOrderSubmitter

Set orderSubmitter = New TradeBuild.SimulatedOrderSubmitter
orderSubmitter.serviceConsumer = serviceConsumer
orderSubmitter.Contract = Contract
orderSubmitter.simulator = mSimulator

serviceConsumer.Ready
Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = orderSubmitter
End Function

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

