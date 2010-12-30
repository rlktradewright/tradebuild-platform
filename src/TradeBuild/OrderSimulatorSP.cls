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

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements TradeBuildSP.ICommonServiceProvider
Implements TradeBuildSP.IOrderSubmissionSrvcProvider

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "OrderSimulatorSP"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mHandle As Long
Private mCommonServiceConsumer As ICommonServiceConsumer

Private mSimulator As AdvancedOrderSimulator

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = "TradeBuild Default Order Simulator"

Set mSimulator = New AdvancedOrderSimulator
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderSimulatorSP terminated"
End Sub

'@================================================================================
' ICommonServiceProvider Interface Members
'@================================================================================

Private Property Get ICommonServiceProvider_Details() As TradeBuildSP.ServiceProviderDetails
Dim details As TradeBuildSP.ServiceProviderDetails
Const ProcName As String = "ICommonServiceProvider_Details"
On Error GoTo Err

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

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub ICommonServiceProvider_Link( _
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, _
                ByVal Handle As Long, _
                ByVal params As Parameters)
Const ProcName As String = "ICommonServiceProvider_Link"
On Error GoTo Err

Set mCommonServiceConsumer = CommonServiceConsumer
mSimulator.CommonServiceConsumer = mCommonServiceConsumer
mHandle = Handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.ExchangeSimulator

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName

End Sub

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)

Const ProcName As String = "ICommonServiceProvider_name"
On Error GoTo Err



Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Property Get ICommonServiceProvider_name() As String
Const ProcName As String = "ICommonServiceProvider_name"
On Error GoTo Err

ICommonServiceProvider_name = name

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Private Sub ICommonServiceProvider_Terminate()
' nothing to do
Const ProcName As String = "ICommonServiceProvider_Terminate"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' IOrderSubmissionSrvcProvider Interface Members
'@================================================================================

Private Function IOrderSubmissionSrvcProvider_CreateOrderSubmitter( _
                ByVal serviceConsumer As TradeBuildSP.IOrderSubmissionSrvcConsumer, _
                ByVal Contract As Contract) As TradeBuildSP.IOrderSubmitter
Dim orderSubmitter As SimulatedOrderSubmitter

Const ProcName As String = "IOrderSubmissionSrvcProvider_CreateOrderSubmitter"

On Error GoTo Err

Set orderSubmitter = New SimulatedOrderSubmitter
orderSubmitter.serviceConsumer = serviceConsumer
orderSubmitter.Contract = Contract
orderSubmitter.simulator = mSimulator

Set IOrderSubmissionSrvcProvider_CreateOrderSubmitter = orderSubmitter

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Private Sub IOrderSubmissionSrvcProvider_StartOrderRecovery( _
                ByVal pOrderRecoveryServiceConsumer As TradeBuildSP.IOrderRecoveryServiceConsumer, _
                ByVal pStartFrom As Date)

Const ProcName As String = "IOrderSubmissionSrvcProvider_StartOrderRecovery"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

