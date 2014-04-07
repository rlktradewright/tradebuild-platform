VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HistDataServiceProvider"
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

Implements IServiceProvider
Implements IHistDataSP

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================


Private Const ModuleName                        As String = "HistDataServiceProvider"

'@===============================================================================
' Enums
'@===============================================================================

'@===============================================================================
' Types
'@===============================================================================

'@===============================================================================
' Member variables
'@===============================================================================

Private mName                                   As String

Private mRole                                   As String

Private mHandle                                 As Long
Private mServiceProviderManager                 As IServiceProviderManager

Private mHistDataStore                          As IHistoricalDataStore

Private mDatabaseType                           As DatabaseTypes
Private mDatabaseName                           As String
Private mServer                                 As String
Private mUserName                               As String
Private mPassword                               As String
Private mUseSynchronousReads                    As Boolean
Private mUseSynchronousWrites                   As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mName = HistoricDataSPName

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""
Role = "INPUT"
mUseSynchronousReads = False
mUseSynchronousWrites = False
End Sub

'@===============================================================================
' IBarDataServiceProvider Interface Members
'@===============================================================================

Private Function IHistDataSP_GetHistoricalDataStore() As IHistoricalDataStore
Const ProcName As String = "IHistDataSP_GetHistoricalDataStore"
On Error GoTo Err

Dim lDbClient As DBClient
Set lDbClient = CreateTradingDBClient(mDatabaseType, mServer, mDatabaseName, mUserName, mPassword, mUseSynchronousReads, mUseSynchronousWrites)

Set mHistDataStore = lDbClient.HistoricalDataStore
Set IHistDataSP_GetHistoricalDataStore = mHistDataStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@===============================================================================
' IServiceProvider Interface Members
'@===============================================================================

Private Property Get IServiceProvider_Details() As ServiceProviderDetails
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
IServiceProvider_Details = details
End Property

Private Sub IServiceProvider_Link( _
                ByVal ServiceProviderManager As IServiceProviderManager, _
                ByVal Handle As Long, _
                ByVal params As Parameters)
Const ProcName As String = "IServiceProvider_Link"
On Error GoTo Err

Set mServiceProviderManager = ServiceProviderManager
mHandle = Handle

Dim param As Parameter
For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameDatabaseType)
        DatabaseType = DatabaseTypeFromString(param.Value)
    Case UCase$(ParamNameDatabaseName)
        DatabaseName = param.Value
    Case UCase$(ParamNameServer)
        Server = param.Value
    Case UCase$(ParamNameUserName)
        Username = param.Value
    Case UCase$(ParamNamePassword)
        Password = param.Value
    Case UCase$(ParamNameUseSynchronousWrites)
        UseSynchronousWrites = gStringToBool(param.Value)
    Case UCase$(ParamNameUseSynchronousReads)
        UseSynchronousReads = gStringToBool(param.Value)
    Case UCase$(ParamNameRole)
        Role = param.Value
    End Select
Next

If mRole = "INPUT" Then
    mServiceProviderManager.RegisterServiceProvider mHandle, _
                                ServiceProviderRoles.SPRoleHistoricalDataInput
Else
    mServiceProviderManager.RegisterServiceProvider mHandle, _
                                ServiceProviderRoles.SPRoleHistoricalDataOutput
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let IServiceProvider_Name(ByVal RHS As String)
Name = RHS
End Property

Private Property Get IServiceProvider_Name() As String
IServiceProvider_Name = Name
End Property

Private Sub IServiceProvider_Terminate()
Const ProcName As String = "IServiceProvider_Terminate"
On Error GoTo Err

mHistDataStore.Finish
Set mServiceProviderManager = Nothing

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' XXXX Event Handlers
'@===============================================================================

'@===============================================================================
' Properties
'@===============================================================================

Public Property Let DatabaseName( _
                ByVal Value As String)
mDatabaseName = Value
End Property

Public Property Let DatabaseType( _
                ByVal Value As DatabaseTypes)
mDatabaseType = Value
End Property

Public Property Get Handle() As Long
Handle = mHandle
End Property

Public Property Let Name(ByVal Value As String)
mName = Value
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Password( _
                ByVal Value As String)
mPassword = Value
End Property
                
Public Property Let Role(ByVal Value As String)
Const ProcName As String = "Role"
On Error GoTo Err

Select Case UCase$(Value)
Case "I", "IN", "INP", "INPUT"
    mRole = "INPUT"
Case "O", "OUT", "OUTPUT"
    mRole = "OUTPUT"
Case Else
    AssertArgument False, "Invalid role parameter"
End Select

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Server( _
                ByVal Value As String)
mServer = Value
End Property
                
Public Property Let Username( _
                ByVal Value As String)
mUserName = Value
End Property

Public Property Let UseSynchronousReads( _
                ByVal Value As Boolean)
mUseSynchronousReads = Value
End Property

Public Property Let UseSynchronousWrites( _
                ByVal Value As Boolean)
mUseSynchronousWrites = Value
End Property

'@===============================================================================
' Methods
'@===============================================================================

'@===============================================================================
' Helper Functions
'@===============================================================================





