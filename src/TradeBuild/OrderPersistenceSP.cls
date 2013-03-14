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

Implements ICommonServiceProvider
Implements IOrderPersistenceServiceProvider

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

Private mCommonServiceConsumer                      As ICommonServiceConsumer

Private mHandle                                     As Long
Private mName                                       As String

Private mRecoveryFilePath                           As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ICommonServiceProvider Interface Members
'@================================================================================

Private Property Get ICommonServiceProvider_Details() As ServiceProviderDetails
Dim details As ServiceProviderDetails
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
                ByVal CommonServiceConsumer As ICommonServiceConsumer, ByVal Handle As Long, ByVal params As Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = Handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.OrderPersistence

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(RecoveryFilePath)
        mRecoveryFilePath = param.Value
    End Select
Next

End Sub

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = mName
End Property

Private Property Let ICommonServiceProvider_name(ByVal Value As String)
mName = Value
End Property

Private Sub ICommonServiceProvider_Terminate()

End Sub

'@================================================================================
' IOrderPersistenceServiceProvider Interface Members
'@================================================================================

Private Property Get IOrderPersistenceServiceProvider_OrderPersistenceDataStore() As IOrderPersistenceDataStore
Const ProcName As String = "IOrderPersistenceServiceProvider_OrderPersistenceDataStore"
On Error GoTo Err

Set IOrderPersistenceServiceProvider_OrderPersistenceDataStore = OrderPersistenceDataStore

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get OrderPersistenceDataStore() As IOrderPersistenceDataStore
Const ProcName As String = "OrderPersistenceDataStore"
On Error GoTo Err

Set OrderPersistenceDataStore = CreateOrderPersistenceDataStore(mRecoveryFilePath)

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


