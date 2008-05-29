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

'@===============================================================================
' Member variables
'@===============================================================================

Private mTradingDB                              As TradingDB
Attribute mTradingDB.VB_VarHelpID = -1
Private mConnectionString                       As String

Private mCommonServiceConsumer                  As ICommonServiceConsumer
Private mHandle                                 As Long
Private mContractInfoServiceConsumer            As IContractInfoServiceConsumer

Private mRequests                               As Collection

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

Set mRequests = New Collection
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
                    "ICommonServiceProvider_Link", _
                    "Invalid role parameter"
        End Select
    End Select
Next

registerSP

createDatabase

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

Private Property Let IContractInfoServiceProvider_ContractInfoServiceConsumer( _
                            ByVal RHS As TradeBuildSP.IContractInfoServiceConsumer)
Set mContractInfoServiceConsumer = RHS
End Property

Private Sub IContractInfoServiceProvider_RequestContractDetails( _
                            ByVal pContractsBuilder As ContractsBuilder)
Dim tc As TaskController
Set tc = mTradingDB.loadContractsAsync(pContractsBuilder.Contracts.ContractSpecifier)

tc.addTaskCompletionListener Me

mRequests.Add pContractsBuilder, tc.taskName

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
' TaskCompletionListener Interface Members
'@===============================================================================

Private Sub TaskCompletionListener_taskCompleted( _
                ByRef ev As TWUtilities30.TaskCompletionEvent)
Dim lBuilder As ContractsBuilder
Dim lContracts As Contracts
Dim lContract As Contract

If ev.cancelled Then
ElseIf ev.errorNumber <> 0 Then
    mCommonServiceConsumer.ServiceProviderError ev.errorNumber, ev.errorMessage, mHandle
Else
    Set lBuilder = mRequests(ev.taskName)
    mRequests.Remove ev.taskName
    If IsEmpty(ev.result) Then
        mContractInfoServiceConsumer.LoadComplete mHandle, lBuilder
    Else
        Set lContracts = ev.result
        For Each lContract In lContracts
            lBuilder.AddContract lContract
        Next
        mContractInfoServiceConsumer.LoadComplete mHandle, lBuilder
    End If
End If
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

Private Sub createDatabase()
If mConnectionString <> "" Then
    Set mTradingDB = CreateTradingDB(mConnectionString)
Else
    Set mTradingDB = CreateTradingDB(GenerateConnectionString(mDatabaseType, _
                                                            mServer, _
                                                            mDatabaseName, _
                                                            mUserName, _
                                                            mPassword))
End If
End Sub

Private Sub handleFatalError( _
                ByVal Number As Long, _
                ByVal Source As String, _
                ByVal Description As String)
mCommonServiceConsumer.FatalServiceProviderError Number, Source, Description, mHandle

Set mCommonServiceConsumer = Nothing
End Sub

Private Sub registerSP()
If mRole = "PRIMARY" Then
    mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                                            ServiceProviderTypes.ContractInfo
Else
    mCommonServiceConsumer.RegisterServiceProvider mHandle, _
                                            ServiceProviderTypes.SecondaryContractInfo
End If
End Sub

