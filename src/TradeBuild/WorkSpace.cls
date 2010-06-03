VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpace"
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

Implements CollectionChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "WorkSpace"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName                       As String

Private mTickers                    As Tickers
Private mOrderPlexes                As OrderPlexes
Private mOrderPlexesSimulated       As OrderPlexes
Private mExecutions                 As Executions
Private mExecutionsSimulated        As Executions

Private mInihibitOrderRecovery      As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTickers = New Tickers
mTickers.AddCollectionChangeListener Me
mTickers.WorkSpace = Me

Set mOrderPlexes = New OrderPlexes
Set mOrderPlexesSimulated = New OrderPlexes

Set mExecutions = New Executions
Set mExecutionsSimulated = New Executions
End Sub

Private Sub Class_Terminate()
Debug.Print "Workspace " & mName & " terminated"
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEvent)
Const ProcName As String = "CollectionChangeListener_Change"

On Error GoTo Err

If TypeOf ev.source Is Tickers Then
    Dim lTicker As Ticker
    Set lTicker = ev.affectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lTicker.PositionManager.OrderContexts.AddCollectionChangeListener Me
        lTicker.PositionManagerSimulated.OrderContexts.AddCollectionChangeListener Me
        lTicker.Executions.AddCollectionChangeListener Me
        lTicker.ExecutionsSimulated.AddCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        ' no need to do anything here
    End Select
ElseIf TypeOf ev.source Is OrderContexts Then
    Dim lOrderContext As OrderContext
    Set lOrderContext = ev.affectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lOrderContext.ActiveOrderPlexes.AddCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        lOrderContext.ActiveOrderPlexes.RemoveCollectionChangeListener Me
    End Select
ElseIf TypeOf ev.source Is OrderPlexes Then
    Dim lOrderPlex As orderPlex
    Set lOrderPlex = ev.affectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        If lOrderPlex.IsSimulated Then
            mOrderPlexesSimulated.Add lOrderPlex
        Else
            mOrderPlexes.Add lOrderPlex
        End If
    Case CollectionChangeTypes.CollItemRemoved
        ' do nothing here because we don't care whether the order plex
        ' is active or not
    End Select
ElseIf TypeOf ev.source Is Executions Then
    Dim lExec As Execution
    Set lExec = ev.affectedItem
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        If lExec.IsSimulated Then
            mExecutionsSimulated.Add lExec
        Else
            mExecutions.Add lExec
        End If
    Case CollectionChangeTypes.CollItemRemoved
        ' should never happen
    End Select
End If

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Executions() As Executions
Set Executions = mExecutions
End Property

Public Property Get ExecutionsSimulated() As Executions
Set ExecutionsSimulated = mExecutionsSimulated
End Property

Public Property Let InihibitOrderRecovery(ByVal value As Boolean)
mInihibitOrderRecovery = value
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Get OrderPlexes() As OrderPlexes
Set OrderPlexes = mOrderPlexes
End Property

Public Property Get OrderPlexesSimulated() As OrderPlexes
Set OrderPlexesSimulated = mOrderPlexesSimulated
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


