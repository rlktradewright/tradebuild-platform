VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpaces"
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

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "WorkSpaces"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWorkspaces                     As EnumerableCollection
Private mCollectionChangeListeners      As Listeners

Private mMarketDataManager              As IMarketDataManager
Private mOrderManager                   As OrderManager
Private mStudyLibraryManager            As StudyLibraryManager
Private mHistoricalDataStore            As IHistoricalDataStore

Private mOrderSubmitterFactoryLive      As IOrderSubmitterFactory
Private mOrderSubmitterFactorySimulated As IOrderSubmitterFactory

Private mMoneyManager                   As IMoneyManager

Private mConfig                         As ConfigurationSection

Private mScopeName                      As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mWorkspaces = New EnumerableCollection
Set mCollectionChangeListeners = New Listeners
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mWorkspaces.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal Name As String) As WorkSpace
Const ProcName As String = "Add"
On Error GoTo Err

Set Add = New WorkSpace
Add.Initialise Name, _
                mScopeName, _
                mMarketDataManager, _
                mHistoricalDataStore, _
                mOrderManager, _
                mOrderSubmitterFactoryLive, _
                mOrderSubmitterFactorySimulated, _
                mMoneyManager
mWorkspaces.Add Add, Name

If Not mConfig Is Nothing Then Add.LoadFromConfig mConfig.AddConfigurationSection(ConfigSectionWorkspace & "(" & Name & ")")

fireCollectionChange CollItemAdded, Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

'gNotifyExistingCollectionMembers mWorkspaces, pListener, Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Contains(ByVal Name As String) As Boolean
Const ProcName As String = "Item"
On Error GoTo Err

Contains = mWorkspaces.Contains(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pScopeName As String, _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pOrderManager As OrderManager, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistoricalDataStore As IHistoricalDataStore, _
                ByVal pOrderSubmitterFactoryLive As IOrderSubmitterFactory, _
                ByVal pOrderSubmitterFactorySimulated As IOrderSubmitterFactory, _
                ByVal pMoneyManager As IMoneyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mOrderManager = pOrderManager
Set mMarketDataManager = pMarketDataManager
Set mStudyLibraryManager = pStudyLibraryManager
Set mHistoricalDataStore = pHistoricalDataStore
Set mOrderSubmitterFactoryLive = pOrderSubmitterFactoryLive
Set mOrderSubmitterFactorySimulated = pOrderSubmitterFactorySimulated
Set mMoneyManager = pMoneyManager
mScopeName = pScopeName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Item(ByVal Name As String) As WorkSpace
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
On Error GoTo Err

Set Item = mWorkspaces.Item(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Assert Not pConfig Is Nothing, "pConfig Is Nothing"

gLog "Loading workspaces from config", ModuleName, ProcName

Set mConfig = pConfig

Dim lCs As ConfigurationSection
For Each lCs In mConfig
    gLog "Loading workspace: " & lCs.InstanceQualifier, ModuleName, ProcName
    Add lCs.InstanceQualifier
Next

gLog "Workspaces loaded from config", ModuleName, ProcName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mWorkspaces.Enumerator
End Function

Public Sub Remove(ByVal pWorkspace As WorkSpace)
Const ProcName As String = "Remove"

On Error GoTo Err

mWorkspaces.Remove pWorkspace.Name
fireCollectionChange CollItemRemoved, pWorkspace

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireCollectionChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As WorkSpace)
Const ProcName As String = "fireCollectionChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mCollectionChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mCollectionChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ICollectionChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




