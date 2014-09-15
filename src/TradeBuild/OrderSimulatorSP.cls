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
Attribute VB_Creatable = True
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

Implements IServiceProvider
Implements IOrderSubmissionSP

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

Private mName                           As String
Private mHandle                         As Long
Private mServiceProviderManager         As IServiceProviderManager
Private mMarketDataManager              As IMarketDataManager
Private mPrimaryContractStore           As IContractStore
Private mSecondaryContractStore         As IContractStore

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mName = "TradeBuild Default Order Simulator"
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderSimulatorSP terminated"
End Sub

'@================================================================================
' IServiceProvider Interface Members
'@================================================================================

Private Property Get IServiceProvider_Details() As ServiceProviderDetails
Dim Details As ServiceProviderDetails
Const ProcName As String = "IServiceProvider_Details"
On Error GoTo Err

With Details
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
IServiceProvider_Details = Details

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IServiceProvider_Link( _
                ByVal ServiceProviderManager As IServiceProviderManager, _
                ByVal Handle As Long, _
                ByVal Params As Parameters, _
                ByVal pSessionID As String)
Const ProcName As String = "IServiceProvider_Link"
On Error GoTo Err

Set mServiceProviderManager = ServiceProviderManager
mHandle = Handle
mServiceProviderManager.RegisterServiceProvider mHandle, _
                            ServiceProviderRoles.SPRoleOrderSubmissionSimulated

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Property Let IServiceProvider_Name(ByVal value As String)

Const ProcName As String = "IServiceProvider_Name"
On Error GoTo Err



Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IServiceProvider_Name() As String
Const ProcName As String = "IServiceProvider_Name"
On Error GoTo Err

IServiceProvider_Name = Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IServiceProvider_Terminate()
' nothing to do
Const ProcName As String = "IServiceProvider_Terminate"
On Error GoTo Err



Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IOrderSubmissionSP Interface Members
'@================================================================================

Private Function IOrderSubmissionSP_GetOrderRecoveryAgent() As IOrderRecoveryAgent

End Function

Private Function IOrderSubmissionSP_GetOrderSubmitterFactory() As IOrderSubmitterFactory
Const ProcName As String = "IOrderSubmissionSP_GetOrderSubmitterFactory"
On Error GoTo Err

Set IOrderSubmissionSP_GetOrderSubmitterFactory = New SimOrderSubmitterFactory

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let Name(ByVal value As String)
mName = value
End Property

Public Property Get Name() As String
Name = mName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pMarketDataManager As IMarketDataManager, ByVal pPrimaryContractStore As IContractStore, ByVal pSecondaryContractStore As IContractStore)
Set mMarketDataManager = pMarketDataManager
Set mPrimaryContractStore = pPrimaryContractStore
Set mSecondaryContractStore = pSecondaryContractStore
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

