VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataSourceSubscriptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type DataSourceSubscription
    DataSource              As MarketDataSource
    Listener                As IGenericTickListener
    SubscribedToMarketData  As Boolean
    SubscribedToMarketDepth As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DataSourceSubscriptions"

'@================================================================================
' Member variables
'@================================================================================

Private mSubscriptions                              As New Collection
Private mNumberOfMarketDataSubscriptions            As Long
Private mNumberOfMarketDepthSubscriptions           As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get NumberOfMarketDataSubscriptions() As Long
NumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions
End Property

Friend Property Get NumberOfMarketDepthSubscriptions() As Long
NumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddMarketDataSubscription(ByVal pDataSource As MarketDataSource, ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddMarketDataSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lKey As String

lKey = getKey(pDataSource, pListener)

On Error Resume Next
lSubscr = mSubscriptions.Item(lKey)
On Error GoTo Err

If Not lSubscr.DataSource Is Nothing Then mSubscriptions.Remove lKey

Set lSubscr.DataSource = pDataSource
Set lSubscr.Listener = pListener
lSubscr.SubscribedToMarketData = True

mSubscriptions.Add lSubscr, lKey

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions + 1
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub AddMarketDepthSubscription( _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "AddMarketDepthSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lKey As String

lKey = getKey(pDataSource, pListener)

On Error Resume Next
lSubscr = mSubscriptions.Item(lKey)
On Error GoTo Err

If Not lSubscr.DataSource Is Nothing Then mSubscriptions.Remove lKey

Set lSubscr.DataSource = pDataSource
Set lSubscr.Listener = pListener
lSubscr.SubscribedToMarketDepth = True

mSubscriptions.Add lSubscr, lKey

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions + 1
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub FireMarketDataError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "FireMarketDataError"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lVar As Variant

For Each lVar In mSubscriptions
    lSubscr = lVar
    If lSubscr.SubscribedToMarketData Then lSubscr.DataSource.NotifyMarketDataError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub FireMarketDepthError( _
                ByVal pErrorCode As Long, _
                ByVal pErrorMsg As String)
Const ProcName As String = "FireMarketDepthError"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lVar As Variant

For Each lVar In mSubscriptions
    lSubscr = lVar
    If lSubscr.SubscribedToMarketDepth Then lSubscr.DataSource.NotifyMarketDepthError pErrorCode, pErrorMsg
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub FireTick( _
                ByVal pStreamId As Long, _
                ByRef pTick As GenericTick)
Const ProcName As String = "FireTick"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lVar As Variant
Dim ev As GenericTickEventData

ev.StreamId = pStreamId
ev.Tick = pTick

For Each lVar In mSubscriptions
    lSubscr = lVar
    Set ev.Source = lSubscr.DataSource
    If pTick.TickType = TickTypeMarketDepth Or _
        pTick.TickType = TickTypeMarketDepthReset _
    Then
        If lSubscr.SubscribedToMarketDepth Then lSubscr.Listener.NotifyTick ev
    Else
        If lSubscr.SubscribedToMarketData Then lSubscr.Listener.NotifyTick ev
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveMarketDataSubscription( _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveMarketDataSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lKey As String

lKey = getKey(pDataSource, pListener)

lSubscr = mSubscriptions.Item(lKey)

lSubscr.SubscribedToMarketData = False

mSubscriptions.Remove lKey
If lSubscr.SubscribedToMarketDepth Then mSubscriptions.Add lSubscr, lKey

mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveMarketDataSubscriptions()
Const ProcName As String = "RemoveMarketDataSubscriptions"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim var As Variant
Dim lKey As String

For Each var In mSubscriptions
    lSubscr = var
    lSubscr.SubscribedToMarketData = False
    lKey = getKey(lSubscr.DataSource, lSubscr.Listener)
    mSubscriptions.Remove lKey
    If lSubscr.SubscribedToMarketDepth Then mSubscriptions.Add lSubscr, lKey
    mNumberOfMarketDataSubscriptions = mNumberOfMarketDataSubscriptions - 1
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscription( _
                ByVal pDataSource As MarketDataSource, _
                ByVal pListener As IGenericTickListener)
Const ProcName As String = "RemoveMarketDepthSubscription"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim lKey As String

lKey = getKey(pDataSource, pListener)

lSubscr = mSubscriptions.Item(lKey)

lSubscr.SubscribedToMarketDepth = False

mSubscriptions.Remove lKey
If lSubscr.SubscribedToMarketData Then mSubscriptions.Add lSubscr, lKey

mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Sub
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Friend Sub RemoveMarketDepthSubscriptions()
Const ProcName As String = "RemoveMarketDepthSubscriptions"
On Error GoTo Err

Dim lSubscr As DataSourceSubscription
Dim var As Variant
Dim lKey As String

For Each var In mSubscriptions
    lSubscr = var
    lSubscr.SubscribedToMarketDepth = False
    lKey = getKey(lSubscr.DataSource, lSubscr.Listener)
    mSubscriptions.Remove lKey
    If lSubscr.SubscribedToMarketData Then mSubscriptions.Add lSubscr, lKey
    mNumberOfMarketDepthSubscriptions = mNumberOfMarketDepthSubscriptions - 1
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getKey(ByVal pDataSource As MarketDataSource, ByVal pListener As IGenericTickListener) As String
getKey = gGetObjectKey(pDataSource) & "/" & gGetObjectKey(pListener)
End Function


