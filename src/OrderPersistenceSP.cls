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
Private mOrderPersistenceServiceConsumer            As IOrderPersistenceServiceConsumer

Private mHandle                                     As Long
Private mName                                       As String

Private mRecoveryFilePath                           As String

Private mTS                                         As TextStream

Private mSessionStarted                             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ICommonServiceProvider Interface Members
'@================================================================================

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
                ByVal CommonServiceConsumer As TradeBuildSP.ICommonServiceConsumer, ByVal Handle As Long, ByVal params As TWUtilities30.Parameters)
Dim param As Parameter

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = Handle
mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                            ServiceProviderTypes.OrderPersistence

For Each param In params
    Select Case UCase$(param.name)
    Case UCase$(RecoveryFilePath)
        mRecoveryFilePath = param.value
    End Select
Next

End Sub

Private Property Get ICommonServiceProvider_name() As String
ICommonServiceProvider_name = mName
End Property

Private Property Let ICommonServiceProvider_name(ByVal RHS As String)
mName = RHS
End Property

Private Sub ICommonServiceProvider_Terminate()
If Not mTS Is Nothing Then mTS.Close
Set mTS = Nothing
End Sub

'@================================================================================
' IOrderPersistenceServiceProvider Interface Members
'@================================================================================

Private Property Let IOrderPersistenceServiceProvider_OrderPersistenceServiceConsumer( _
                ByVal RHS As TradeBuildSP.IOrderPersistenceServiceConsumer)
Set mOrderPersistenceServiceConsumer = RHS
End Property

Private Sub IOrderPersistenceServiceProvider_RequestOrderInfo()
Dim fso As New FileSystemObject

Const ProcName As String = "IOrderPersistenceServiceProvider_RequestOrderInfo"

On Error GoTo Err

On Error Resume Next
Set mTS = fso.OpenTextFile(mRecoveryFilePath, IOMode.ForReading, False)
On Error GoTo Err

Do While sendNextOrderInfo
Loop

mOrderPersistenceServiceConsumer.PersistedOrderInfoEnd

Exit Sub

Err:
HandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName
mCommonServiceConsumer.NotifyFatalError Err.number, ProcName, Err.Description, mHandle
End Sub

Private Sub IOrderPersistenceServiceProvider_SaveOrderInfo( _
                orderInfo As String)
Const ProcName As String = "IOrderPersistenceServiceProvider_SaveOrderInfo"
On Error GoTo Err

If Not mSessionStarted Then Exit Sub

mTS.writeLine orderInfo

Exit Sub

Err:
HandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName
mCommonServiceConsumer.NotifyFatalError Err.number, ProcName, Err.Description, mHandle
End Sub

Private Sub IOrderPersistenceServiceProvider_StartSession()
Const ProcName As String = "IOrderPersistenceServiceProvider_StartSession"
Dim fso As New FileSystemObject

On Error GoTo Err

If Not mTS Is Nothing Then mTS.Close

Set mTS = fso.CreateTextFile(mRecoveryFilePath, True)
mTS.writeLine "//Session started at " & FormatTimestamp(Now, TimestampDateAndTimeISO8601)

mSessionStarted = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=False, pLog:=True, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName
mCommonServiceConsumer.NotifyFatalError Err.number, ProcName, Err.Description, mHandle
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getNextLine(ByRef result As String) As Boolean
Const ProcName As String = "getNextLine"
Dim failpoint As String
On Error GoTo Err

Do While Not mTS.AtEndOfStream
    result = mTS.readLine
    If result <> "" And Left$(result, 2) <> "//" Then
        getNextLine = True
        Exit Do
    End If
Loop

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Private Function sendNextOrderInfo() As Boolean
Dim orderInfo As String
Const ProcName As String = "sendNextOrderInfo"
Dim failpoint As String
On Error GoTo Err

If getNextLine(orderInfo) Then
    mOrderPersistenceServiceConsumer.PersistedOrderInfo orderInfo
    sendNextOrderInfo = True
End If

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function
