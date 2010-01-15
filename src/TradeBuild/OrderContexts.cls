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
                ev As CollectionChangeEvent)

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

Private mOrderContexts As Collection
Private mChangeListeners As Collection

Private mEnabled As Boolean

Private mIsSimulated As Boolean

Private mTicker As Ticker
Private mPositionManager As PositionManager

Private mOrderSubmitter As TradeBuildSP.IOrderSubmitter

Private mProviderReady As Boolean

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
Dim failpoint As Long
On Error GoTo Err

Count = mOrderContexts.Count

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get DefaultOrderContext() As OrderContext
Set DefaultOrderContext = Item(DefaultOrderContextName)
End Property

Friend Property Let Enabled(ByVal value As Boolean)
Dim oc As OrderContext
Const ProcName As String = "Enabled"
Dim failpoint As Long
On Error GoTo Err

mEnabled = value
For Each oc In mOrderContexts
    oc.Enabled = mEnabled
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

Friend Property Let IsSimulated(ByVal value As Boolean)
Dim oc As OrderContext
Const ProcName As String = "IsSimulated"
Dim failpoint As Long
On Error GoTo Err

mIsSimulated = value
For Each oc In mOrderContexts
    oc.IsSimulated = mIsSimulated
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

Friend Property Let orderSubmitter( _
                ByVal value As TradeBuildSP.IOrderSubmitter)
Dim lOrderContext As OrderContext
Const ProcName As String = "OrderSubmitter"
Dim failpoint As Long
On Error GoTo Err

Set mOrderSubmitter = value
For Each lOrderContext In mOrderContexts
    lOrderContext.orderSubmitter = value
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let ProviderReady(ByVal value As Boolean)
Dim lOrderContext As OrderContext
Const ProcName As String = "ProviderReady"
Dim failpoint As Long
On Error GoTo Err

mProviderReady = value
For Each lOrderContext In mOrderContexts
    lOrderContext.IsReady = value
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

Friend Property Let Ticker(ByVal value As Ticker)
Dim lOrderContext As OrderContext
Const ProcName As String = "Ticker"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = value
For Each lOrderContext In mOrderContexts
    lOrderContext.Ticker = value
Next

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(ByVal name As String) As OrderContext
Const ProcName As String = "Add"
Dim failpoint As Long
On Error GoTo Err

Set Add = New OrderContext
Add.Ticker = mTicker
Add.PositionManager = mPositionManager
Add.orderSubmitter = mOrderSubmitter
Add.IsSimulated = mIsSimulated
Add.name = name
mOrderContexts.Add Add, name
fireChange CollItemAdded, Add
Add.Enabled = mEnabled
Add.IsReady = mProviderReady

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)

Const ProcName As String = "AddCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

mChangeListeners.Add value

gNotifyExistingCollectionMembers mOrderContexts, value, Me

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Finish()
Dim OrderContext As OrderContext

Const ProcName As String = "Finish"
Dim failpoint As Long
On Error GoTo Err

Set mTicker = Nothing
Set mPositionManager = Nothing
Set mOrderSubmitter = Nothing
Set mChangeListeners = Nothing

For Each OrderContext In mOrderContexts
    OrderContext.Finish
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Friend Sub Initialise(ByVal pPositionManager As PositionManager)
Set mPositionManager = pPositionManager
Add DefaultOrderContextName
End Sub

Public Function Item(index As Variant) As OrderContext
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As Long
On Error GoTo Err

Set Item = mOrderContexts(index)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As Long
On Error GoTo Err

Set NewEnum = mOrderContexts.[_NewEnum]

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Function

Public Sub Remove( _
                ByVal pOrderContext As OrderContext)
Const ProcName As String = "Remove"
Dim failpoint As Long
On Error GoTo Err

mOrderContexts.Remove pOrderContext.name
fireChange CollItemRemoved, pOrderContext

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As Long
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As OrderContext)
Dim listener As CollectionChangeListener
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As Long
On Error GoTo Err

Set ev.source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For Each listener In mChangeListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pNumber:=Err.number, pSource:=Err.source, pDescription:=Err.Description, pProjectName:=ProjectName, pModuleName:=ModuleName, pFailpoint:=failpoint
End Sub


