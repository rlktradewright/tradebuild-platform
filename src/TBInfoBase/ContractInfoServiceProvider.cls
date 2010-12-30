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
Implements TaskCompletionListener

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================

Private Const ProjectName                       As String = "TBInfoBase26"
Private Const ModuleName                        As String = "ContractInfoSrvcProvider"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

Private Type ContractRequest
    serviceConsumer     As IContractInfoServiceConsumer
    builder             As ContractsBuilder
    numberOfRequests    As Long
    next                As Long
End Type

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                              As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString                       As String

Private mCommonServiceConsumer                  As ICommonServiceConsumer
Private mHandle                                 As Long

Private mRequests()                             As ContractRequest
Private mRequestsMaxIndex                       As Long
Private mFirstFreeRequestIndex                  As Long

Private mName                                   As String

Private mDatabaseType                           As DatabaseTypes
Private mDatabaseName                           As String
Private mServer                                 As String
Private mUserName                               As String
Private mPassword                               As String
Private mRole                                   As String

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mName = ContractInfoSPName

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""
mRole = "PRIMARY"

ReDim mRequests(3) As ContractRequest
mRequestsMaxIndex = -1
mFirstFreeRequestIndex = -1
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
Const ProcName As String = "ICommonServiceProvider_Link"
Dim param As Parameter

Dim failpoint As Long
On Error GoTo Err

Set mCommonServiceConsumer = CommonServiceConsumer
mHandle = handle

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
    Case UCase$(ParamNameRole)
        Select Case UCase$(param.value)
        Case "P", "PR", "PRIM", "PRIMARY"
            mRole = "PRIMARY"
        Case "S", "SEC", "SECOND", "SECONDARY"
            mRole = "SECONDARY"
        Case Else
            Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                    ProjectName & "." & ModuleName & ":" & ProcName, _
                    "Invalid role parameter"
        End Select
    End Select
Next

registerSP

createDatabase

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

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

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                            ByVal serviceConsumer As TradeBuildSP.IContractInfoServiceConsumer, _
                            ByVal pContractsBuilder As ContractsBuilder)
Const ProcName As String = "IContractInfoServiceProvider_RequestContractDetails"
Dim contractSpec As ContractSpecifier
Dim tc As TaskController
Dim index As Long

Dim failpoint As Long
On Error GoTo Err

index = allocateRequestEntry
Set mRequests(index).serviceConsumer = serviceConsumer
Set mRequests(index).builder = pContractsBuilder

Set contractSpec = pContractsBuilder.Contracts.ContractSpecifier

logMessage pMsg:="Requesting contract data for " & contractSpec.ToString, pProcName:="IContractInfoServiceProvider_RequestContractDetails"

If contractSpec.SecType = SecTypeNone Then
    ' no sectype supplied, so we'll place a separate request for each of the following:
    '       Futures
    '       Indexes
    '       Stocks
    Set tc = mTradingDB.LoadContractsAsync(CreateContractSpecifier(contractSpec.LocalSymbol, _
                                                    contractSpec.Symbol, _
                                                    contractSpec.Exchange, _
                                                    SecTypeFuture, _
                                                    contractSpec.CurrencyCode, _
                                                    contractSpec.Expiry), _
                                            index)
    tc.AddTaskCompletionListener Me
    incrementrequestcount index
    
    If contractSpec.LocalSymbol = "" Then
        Set tc = mTradingDB.LoadContractsAsync(CreateContractSpecifier(contractSpec.Symbol, _
                                                        "", _
                                                        contractSpec.Exchange, _
                                                        SecTypeFuture, _
                                                        contractSpec.CurrencyCode, _
                                                        contractSpec.Expiry), _
                                                index)
        tc.AddTaskCompletionListener Me
        incrementrequestcount index
    End If
    
    Set tc = mTradingDB.LoadContractsAsync(CreateContractSpecifier(contractSpec.LocalSymbol, _
                                                    contractSpec.Symbol, _
                                                    contractSpec.Exchange, _
                                                    SecTypeIndex, _
                                                    contractSpec.CurrencyCode, _
                                                    contractSpec.Expiry), _
                                            index)
    tc.AddTaskCompletionListener Me
    incrementrequestcount index

    Set tc = mTradingDB.LoadContractsAsync(CreateContractSpecifier(contractSpec.LocalSymbol, _
                                                    contractSpec.Symbol, _
                                                    contractSpec.Exchange, _
                                                    SecTypeStock, _
                                                    contractSpec.CurrencyCode, _
                                                    contractSpec.Expiry), _
                                            index)
    tc.AddTaskCompletionListener Me
    incrementrequestcount index

    Set tc = mTradingDB.LoadContractsAsync(CreateContractSpecifier(contractSpec.LocalSymbol, _
                                                    contractSpec.Symbol, _
                                                    contractSpec.Exchange, _
                                                    SecTypeCash, _
                                                    contractSpec.CurrencyCode, _
                                                    contractSpec.Expiry), _
                                            index)
    tc.AddTaskCompletionListener Me
    incrementrequestcount index

Else
    Set tc = mTradingDB.LoadContractsAsync(contractSpec, _
                                            index)
    tc.AddTaskCompletionListener Me
    incrementrequestcount index
End If

Exit Sub

Err:
handleFatalError pProcName:="IContractInfoServiceProvider_RequestContractDetails"
End Sub

Private Sub IContractInfoServiceProvider_StoreContractDetails( _
                            ByVal serviceConsumer As TradeBuildSP.IContractInfoServiceConsumer, _
                            ByVal pContract As Contract)

End Sub

Private Function IContractInfoServiceProvider_Supports( _
                            ByVal capabilities As Long) As Boolean
Dim supportedCapabilities As Long

supportedCapabilities = ContractInfoServiceProviderCapabilities.Store
IContractInfoServiceProvider_Supports = supportedCapabilities And capabilities

End Function

'@===============================================================================
' TaskCompletionListener Interface Members
'@===============================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ByRef ev As TWUtilities30.TaskCompletionEventData)
Const ProcName As String = "TaskCompletionListener_taskCompleted"
Dim lContracts As Contracts
Dim lContract As Contract
Dim index As Long


Dim failpoint As Long
On Error GoTo Err

index = CLng(ev.cookie)
If ev.Cancelled Then
ElseIf ev.ErrorNumber <> 0 Then
    decrementRequestCount index
    mCommonServiceConsumer.NotifyRecoverableError ev.ErrorNumber, ev.ErrorMessage, mHandle
Else
    If Not IsEmpty(ev.result) Then
        Set lContracts = ev.result
        For Each lContract In lContracts
            mRequests(index).builder.AddContract lContract
        Next
    End If
    If decrementRequestCount(index) = 0 Then
        logMessage pMsg:="received contract data for " & _
                        mRequests(index).builder.Contracts.ContractSpecifier.ToString, _
                        pProcName:="TaskCompletionListener_taskCompleted"
        mRequests(index).serviceConsumer.LoadComplete mHandle
        releaseRequestEntry index
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Function allocateRequestEntry() As Long
Const ProcName As String = "allocateRequestEntry"
On Error GoTo Err

If mFirstFreeRequestIndex <> -1 Then
    allocateRequestEntry = mFirstFreeRequestIndex
    mFirstFreeRequestIndex = mRequests(mFirstFreeRequestIndex).next
Else
    mRequestsMaxIndex = mRequestsMaxIndex + 1
    If mRequestsMaxIndex > UBound(mRequests) Then ReDim Preserve mRequests(2 * (UBound(mRequests) + 1) - 1) As ContractRequest
    allocateRequestEntry = mRequestsMaxIndex
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub createDatabase()
Const ProcName As String = "createDatabase"
Dim failpoint As Long
On Error GoTo Err

If mConnectionString <> "" Then
    Set mTradingDB = CreateTradingDB(mConnectionString, mDatabaseType)
Else
    Set mTradingDB = CreateTradingDB(GenerateConnectionString(mDatabaseType, _
                                                            mServer, _
                                                            mDatabaseName, _
                                                            mUserName, _
                                                            mPassword), _
                                    mDatabaseType)
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function decrementRequestCount( _
                ByVal index As Long) As Long
Const ProcName As String = "decrementRequestCount"
On Error GoTo Err

mRequests(index).numberOfRequests = mRequests(index).numberOfRequests - 1
decrementRequestCount = mRequests(index).numberOfRequests

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub handleFatalError( _
                ByRef pProcName As String, _
                Optional ByVal pFailpoint As String)
Dim errNum As Long: errNum = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDesc As String: errDesc = Err.Description

On Error GoTo Err

' re-raise the error to get the calling procedure's procName into the source info
errSource = errSource & vbCrLf & _
            ProjectName & "." & _
            ModuleName & ":" & _
            pProcName & _
            IIf(pFailpoint <> "", " At " & pFailpoint, "")

Err.Raise errNum, errSource, errDesc

' NB: will never get to here so no need for Exit Sub

Err:
mCommonServiceConsumer.NotifyFatalError Err.Number, Err.Source, Err.Description, mHandle
End Sub

Private Sub incrementrequestcount( _
                ByVal index As Long)
Const ProcName As String = "incrementrequestcount"
On Error GoTo Err

mRequests(index).numberOfRequests = mRequests(index).numberOfRequests + 1

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub logMessage( _
                ByRef pMsg As String, _
                ByRef pProcName As String, _
                Optional ByRef pMsgQualifier As String = vbNullString, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal)
gLog pMsg:=pMsg, pMsgQualifier:=pMsgQualifier, pProcName:=pProcName, pProjName:=ProjectName, pModName:=ModuleName, pLogLevel:=pLogLevel
End Sub

Private Sub registerSP()
Const ProcName As String = "registerSP"
Dim failpoint As Long
On Error GoTo Err

If mRole = "PRIMARY" Then
    mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                                            ServiceProviderTypes.ContractInfo
Else
    mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                                            ServiceProviderTypes.SecondaryContractInfo
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub releaseRequestEntry( _
                ByVal index As Long)
Const ProcName As String = "releaseRequestEntry"
On Error GoTo Err

mRequests(index).next = mFirstFreeRequestIndex
mFirstFreeRequestIndex = index
Set mRequests(index).builder = Nothing
mRequests(index).numberOfRequests = 0

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

