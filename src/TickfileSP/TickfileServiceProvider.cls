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

Implements IServiceProvider
Implements ITickfileSP

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "TickfileServiceProvider"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mName                                   As String

Private mRole                                   As String

Private mHandle                                 As Long
Private mServiceProviderManager                 As IServiceProviderManager

Private mTickfileDataReaders                    As Collection
Private mTickfileDataWriters                    As Collection

Private mTickfilePath                           As String

Private mTickfileGranularity                    As TickfileGranularities

Private mTickfileStore                          As ITickfileStore

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Set mTickfileDataReaders = New Collection
Set mTickfileDataWriters = New Collection

mName = App.ProductName
Role = "INPUT"
TickfileGranularity = "FILE PER WEEK"
End Sub

'================================================================================
' IServiceProvider Interface Members
'================================================================================

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
                ByVal pServiceProviderManager As IServiceProviderManager, _
                ByVal pHandle As Long, _
                ByVal pParams As Parameters, _
                ByVal pSessionID As String)
Const ProcName As String = "IServiceProvider_Link"
On Error GoTo Err

Set mServiceProviderManager = pServiceProviderManager
mHandle = pHandle

Dim param As Parameter
For Each param In pParams
    Select Case UCase$(param.Name)
    Case UCase$(ParamNameTickfilePath)
        mTickfilePath = param.Value
    Case UCase$(ParamNameRole)
        Role = param.Value
    Case UCase$(ParamNameTickfileGranularity)
        Select Case UCase$(param.Value)
        Case "FILE PER DAY", "DAY"
            mTickfileGranularity = TickfileGranularityDay
        Case "FILE PER WEEK", "WEEK"
            mTickfileGranularity = TickfileGranularityWeek
        Case "FILE PER TRADING SESSION", "FILE PER SESSION", "SESSION"
            mTickfileGranularity = TickfileGranularitySession
        Case "FILE PER EXECUTION", "EXECUTION"
            mTickfileGranularity = TickfileGranularityExecution
        Case Else
            AssertArgument False, "Invalid tickfile granularity parameter"
            Exit Sub
        End Select
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
Set mServiceProviderManager = Nothing

Exit Sub
Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'================================================================================
' ITickfileSP Interface Members
'================================================================================

Private Function ITickfileSP_GetTickfileStore() As ITickfileStore
Const ProcName As String = "ITickfileSP_GetTickfileStore"
On Error GoTo Err

If mRole = "INPUT" Then
    Set mTickfileStore = CreateTickfileStore(TickfileReadOnly, mTickfilePath, mTickfileGranularity)
Else
    Set mTickfileStore = CreateTickfileStore(TickfileWriteOnly, mTickfilePath, mTickfileGranularity)
End If
Set ITickfileSP_GetTickfileStore = mTickfileStore

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get Handle() As Long
Handle = mHandle
End Property

Friend Property Let Name(ByVal Value As String)
mName = Value
End Property

Public Property Get Name() As String
Name = mName
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

Public Property Let TickfileGranularity(ByVal Value As String)
Const ProcName As String = "TickfileGranularity"
On Error GoTo Err

Select Case UCase$(Value)
Case "FILE PER DAY", "DAY"
    mTickfileGranularity = TickfileGranularityDay
Case "FILE PER WEEK", "WEEK"
    mTickfileGranularity = TickfileGranularityWeek
Case "FILE PER TRADING SESSION", "FILE PER SESSION", "SESSION"
    mTickfileGranularity = TickfileGranularitySession
Case "FILE PER EXECUTION", "EXECUTION"
    mTickfileGranularity = TickfileGranularityExecution
Case Else
    AssertArgument False, "Invalid tickfile granularity parameter"
End Select

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================


