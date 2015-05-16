VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileServiceProvider"
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
Implements ITickfileSP

'@===============================================================================
' Events
'@===============================================================================

'@===============================================================================
' Constants
'@===============================================================================


Private Const ModuleName                        As String = "TickfileServiceProvider"

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

Private mTickfileStore                          As ITickfileStore

Private mDatabaseType                           As DatabaseTypes
Private mDatabaseName                           As String
Private mServer                                 As String
Private mUserName                               As String
Private mPassword                               As String
Private mUseSynchronousWrites                   As Boolean
Private mUseSynchronousReads                    As Boolean

'@===============================================================================
' Class Event Handlers
'@===============================================================================

Private Sub Class_Initialize()
mName = SQLDBTickfileSPName

mDatabaseType = DbSQLServer2005
mDatabaseName = "Trading"
mServer = "localhost"
mUserName = ""
mPassword = ""

Role = "INPUT"
End Sub

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
                ByVal params As Parameters, _
                ByVal pSessionID As String)
Const ProcName As String = "IServiceProvider_Link"
On Error GoTo Err

Set mServiceProviderManager = ServiceProviderManager
mHandle = Handle

Dim param As Parameter
For Each param In params
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameDatabaseType)
        DatabaseType = DatabaseTypeFromString(param.Value)
        AssertArgument DatabaseType <> DbNone, "Invalid database type parameter"
    Case UCase$(ParamNameDatabaseName)
        DatabaseName = param.Value
    Case UCase$(ParamNameServer)
        Server = param.Value
    Case UCase$(ParamNameUserName)
        Username = param.Value
    Case UCase$(ParamNamePassword)
        Password = param.Value
    Case UCase$(ParamNameUseSynchronousReads)
        useSynchronousReads = gStringToBool(param.Value)
    Case UCase$(ParamNameUseSynchronousWrites)
        UseSynchronousWrites = gStringToBool(param.Value)
    Case UCase$(ParamNameRole)
        Role = param.Value
    End Select
Next

If mRole = "INPUT" Then
    mServiceProviderManager.RegisterServiceProvider mHandle, _
                                ServiceProviderRoles.SPRoleTickfileInput
Else
    mServiceProviderManager.RegisterServiceProvider mHandle, _
                                ServiceProviderRoles.SPRoleTickfileOutput
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

mTickfileStore.Finish

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@===============================================================================
' ITickfileServiceProvider Interface Members
'@===============================================================================

Private Function ITickfileSP_GetTickfileStore() As ITickfileStore
Const ProcName As String = "ITickfileSP_GetTickfileStore"
On Error GoTo Err

Dim lDbClient As DBClient
Set lDbClient = CreateTradingDBClient(mDatabaseType, mServer, mDatabaseName, mUserName, mPassword)

Set mTickfileStore = lDbClient.TickfileStore
Set ITickfileSP_GetTickfileStore = mTickfileStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@===============================================================================
' xxxx Event Handlers
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

Friend Property Get DatabaseType() As DatabaseTypes
DatabaseType = mDatabaseType
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

Public Property Let UseSynchronousWrites( _
                ByVal Value As Boolean)
mUseSynchronousWrites = Value
End Property

Public Property Let useSynchronousReads( _
                ByVal Value As Boolean)
mUseSynchronousReads = Value
End Property

'@===============================================================================
' Methods
'@===============================================================================

'@===============================================================================
' Helper Functions
'@===============================================================================


