VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OrderContexts"
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

Private Const ModuleName                    As String = "OrderContexts"

Private Const DefaultOrderContextName       As String = "$default"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mDataSource                         As IMarketDataSource
Private mOrderSubmitter                     As IOrderSubmitter
Private mOrderAuthoriser                    As IOrderAuthoriser
Private mAccumulatedBracketOrders           As Orders
Private mAccumulatedOrders                  As Orders

Private mOrderContexts                      As Collection
Private mChangeListeners                    As Collection

Private mEnabled                            As Boolean

Private mIsSimulated                        As Boolean

Private mProviderReady                      As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mOrderContexts = New Collection
Set mChangeListeners = New Collection
mEnabled = True
End Sub

Private Sub Class_Terminate()
Debug.Print "OrderContexts terminated"
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
Const ProcName As String = "Count"

On Error GoTo Err

Count = mOrderContexts.Count

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DefaultOrderContext() As OrderContext
Set DefaultOrderContext = Item(DefaultOrderContextName)
End Property

Friend Property Let Enabled(ByVal value As Boolean)
Dim oc As OrderContext
Const ProcName As String = "Enabled"

On Error GoTo Err

mEnabled = value
For Each oc In mOrderContexts
    oc.Enabled = mEnabled
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let ProviderReady(ByVal value As Boolean)
Dim lOrderContext As OrderContext
Const ProcName As String = "ProviderReady"

On Error GoTo Err

mProviderReady = value
For Each lOrderContext In mOrderContexts
    lOrderContext.IsReady = value
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let DataSource(ByVal value As IMarketDataSource)
Dim lOrderContext As OrderContext
Const ProcName As String = "DataSource"

On Error GoTo Err

Set mDataSource = value
For Each lOrderContext In mOrderContexts
    lOrderContext.Ticker = value
Next

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal pName As String) As OrderContext
Const ProcName As String = "Add"

On Error GoTo Err

Set Add = New OrderContext
Add.Initialise pName, mIsSimulated, mDataSource, mOrderSubmitter, mOrderAuthoriser, mAccumulatedBracketOrders, mAccumulatedOrders
 
mOrderContexts.Add Add, pName
fireChange CollItemAdded, Add
Add.Enabled = mEnabled
Add.IsReady = mProviderReady

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)

Const ProcName As String = "AddCollectionChangeListener"

On Error GoTo Err

mChangeListeners.Add value

gNotifyExistingCollectionMembers mOrderContexts, value, Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Finish()
Dim OrderContext As OrderContext

Const ProcName As String = "Finish"

On Error GoTo Err

Set mDataSource = Nothing
Set mPositionManager = Nothing
Set mOrderSubmitter = Nothing
Set mChangeListeners = Nothing

For Each OrderContext In mOrderContexts
    OrderContext.Finish
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pIsSimulated As Boolean, _
                ByVal pDataSource As IMarketDataSource, _
                ByVal pOrderSubmitter As IOrderSubmitter, _
                ByVal pOrderAuthoriser As IOrderAuthoriser, _
                ByVal pAccumulatedBracketOrders As Orders, _
                ByVal pAccumulatedOrders As Orders)
mIsSimulated = pIsSimulated
Set mDataSource = pDataSource
Set mOrderSubmitter = pOrderSubmitter
Set pOrderAuthoriser = pOrderAuthoriser
Set mAccumulatedBracketOrders = pAccumulatedBracketOrders
Set mAccumulatedOrders = pAccumulatedOrders
Add DefaultOrderContextName
End Sub

Public Function Item(index As Variant) As OrderContext
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mOrderContexts(index)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"

On Error GoTo Err

Set NewEnum = mOrderContexts.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub Remove( _
                ByVal pOrderContext As OrderContext)
Const ProcName As String = "Remove"

On Error GoTo Err

mOrderContexts.Remove pOrderContext.Name
fireChange CollItemRemoved, pOrderContext

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"

On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As OrderContext)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEventData
Const ProcName As String = "fireChange"

On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


