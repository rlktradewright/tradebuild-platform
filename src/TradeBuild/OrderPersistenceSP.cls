VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderPersistenceSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IServiceProvider
Implements IOrderPersistenceSP

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

Private Const ModuleName                            As String = "OrderPersistenceSP"

Private Const RecoveryFilePath                      As String = "RecoveryFilePath"

'@================================================================================
' Member variables
'@================================================================================

Private mServiceProviderManager                     As IServiceProviderManager

Private mHandle                                     As Long
Private mName                                       As String

Private mRecoveryFilePath                           As String

Private mOrderPersistenceDataStore                  As IOrderPersistenceDataStore

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mRecoveryFilePath = ApplicationSettingsFolder
End Sub

'@================================================================================
' IServiceProvider Interface Members
'@================================================================================

Private Property Get IServiceProvider_Details() As ServiceProviderDetails
Dim Details As ServiceProviderDetails
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
End Property

Private Sub IServiceProvider_Link( _
                ByVal ServiceProviderManager As IServiceProviderManager, _
                ByVal Handle As Long, _
                ByVal Params As Parameters, _
                ByVal pSessionID As String)
Dim param As Parameter

Set mServiceProviderManager = ServiceProviderManager
mHandle = Handle
mServiceProviderManager.RegisterServiceProvider mHandle, _
                            ServiceProviderRoles.SPRoleOrderPersistence

For Each param In Params
    Select Case UCase$(param.Name)
    Case UCase$(RecoveryFilePath)
        If param.Value <> "" Then mRecoveryFilePath = param.Value
    End Select
Next

End Sub

Private Property Get IServiceProvider_Name() As String
IServiceProvider_Name = mName
End Property

Private Property Let IServiceProvider_Name(ByVal Value As String)
mName = Value
End Property

Private Sub IServiceProvider_Terminate()
If Not mOrderPersistenceDataStore Is Nothing Then mOrderPersistenceDataStore.Finish
Set mOrderPersistenceDataStore = Nothing
End Sub

'@================================================================================
' IOrderPersistenceSP Interface Members
'@================================================================================

Private Function IOrderPersistenceSP_GetOrderPersistenceDataStore() As IOrderPersistenceDataStore
Const ProcName As String = "IOrderPersistenceSP_GetOrderPersistenceDataStore"
On Error GoTo Err

Set IOrderPersistenceSP_GetOrderPersistenceDataStore = OrderPersistenceDataStore

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

Public Property Get OrderPersistenceDataStore() As IOrderPersistenceDataStore
Const ProcName As String = "OrderPersistenceDataStore"
On Error GoTo Err

If mOrderPersistenceDataStore Is Nothing Then Set mOrderPersistenceDataStore = createOrderPersistenceDataStore(mRecoveryFilePath)
Set OrderPersistenceDataStore = mOrderPersistenceDataStore

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


