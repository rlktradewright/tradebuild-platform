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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTradeBuildAPIRef           As WeakReference

Private mName                       As String

Private mTickers                    As Tickers
Private mOrderPlexes                As orderPlexes
Private mExecutions                 As executions

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTickers = New Tickers
mTickers.addCollectionChangeListener Me
Set mOrderPlexes = New orderPlexes
Set mExecutions = New executions
End Sub

Private Sub Class_Terminate()
Debug.Print "Workspace " & mName & " terminated"
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change(ev As CollectionChangeEvent)
If TypeOf ev.affectedObject Is Ticker Then
    Dim lTicker As Ticker
    Set lTicker = ev.affectedObject
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lTicker.PositionManager.OrderContexts.addCollectionChangeListener Me
        lTicker.executions.addCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        ' no need to do anything here
    End Select
ElseIf TypeOf ev.affectedObject Is OrderContext Then
    Dim lOrderContext As OrderContext
    Set lOrderContext = ev.affectedObject
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        lOrderContext.activeOrderPlexes.addCollectionChangeListener Me
    Case CollectionChangeTypes.CollItemRemoved
        lOrderContext.activeOrderPlexes.removeCollectionChangeListener Me
    End Select
ElseIf TypeOf ev.affectedObject Is orderPlex Then
    Dim lOrderPlex As orderPlex
    Set lOrderPlex = ev.affectedObject
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        mOrderPlexes.add lOrderPlex
    Case CollectionChangeTypes.CollItemRemoved
        ' do nothing here because we don't care whether the order plex
        ' is active or not
    End Select
ElseIf TypeOf ev.affectedObject Is Execution Then
    Dim lExec As Execution
    Set lExec = ev.affectedObject
    Select Case ev.changeType
    Case CollectionChangeTypes.CollItemAdded
        mExecutions.add lExec
    Case CollectionChangeTypes.CollItemRemoved
        ' should never happen
        tb.notifyFatalError ErrorCodes.ErrUnsupportedOperationException, _
                                        "Unexpected execution deletion", _
                                        "TradeBuild25.Workspace::CollectionChangeListener_Change"
    End Select
End If
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get executions() As executions
Set executions = mExecutions
End Property

Public Property Let name(ByVal value As String)
mName = value
End Property

Public Property Get name() As String
name = mName
End Property

Public Property Get orderPlexes() As orderPlexes
Set orderPlexes = mOrderPlexes
End Property

Public Property Get Tickers() As Tickers
Set Tickers = mTickers
End Property

Friend Property Let tradeBuildAPI( _
                ByVal value As tradeBuildAPI)
Set mTradeBuildAPIRef = weakrefs.createWeakReference(value)
mTickers.tradeBuildAPI = value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function tb() As tradeBuildAPI
Set tb = mTradeBuildAPIRef.Target
End Function

