VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MockDataSource"
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

Implements IMarketDataSource

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "MockDataSource"

'@================================================================================
' Member variables
'@================================================================================

Private mGenericTickListeners                       As New Listeners

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1
Private mContractSpec                               As IContractSpecifier

Private mContractFuture                             As IFuture

Private mTicks()                                    As GenericTick

Private mMktDataManagerRef                          As WeakReference

Private mStreamId                                   As Long

Private mMarketDataRequested                        As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IMarketDataSource Interface Members
'@================================================================================

Private Property Get IMarketDataSource_ClockFuture() As IFuture
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_ContractFuture() As IFuture
Set IMarketDataSource_ContractFuture = mContractFuture
End Property

Private Property Get IMarketDataSource_CurrentQuote(ByVal pTickType As TickTypes) As Quote
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_CurrentTick(ByVal pTickType As TickTypes) As GenericTick
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_ErrorMessage() As String
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_Handle() As Long
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_HasCurrentTick(ByVal pTickType As TickTypes) As Boolean
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_IsDataDelayed() As Boolean

End Property

Private Property Get IMarketDataSource_IsFinished() As Boolean
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_IsMarketDataRequested() As Boolean
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_IsMarketDepthRequested() As Boolean
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_IsTickReplay() As Boolean
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_NumberOfMarketDepthRows() As Long
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_Options() As MarketDataSourceOptions
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Let PriorityTickListener(ByVal Value As IGenericTickListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_Key() As String
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_PriceChange() As PriceChange
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Let IMarketDataSource_PriorityTickListener(ByVal Value As IGenericTickListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_RecoveryContractSpec() As ContractUtils27.IContractSpecifier
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_SessionFuture() As IFuture
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_State() As MarketDataSourceStates
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_StudyBase() As IStudyBase
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Property Get IMarketDataSource_Timestamp() As Date
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Property

Private Sub IMarketDataSource_AddErrorListener(ByVal pListener As IErrorListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddGenericTickListener(ByVal pListener As IGenericTickListener)
mGenericTickListeners.Add pListener
End Sub

Private Sub IMarketDataSource_AddMarketDepthListener(ByVal pListener As IMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddOptionComputationListener(ByVal pListener As IOptionComputationListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddPriceChangeListener(ByVal pListener As IPriceChangeListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddQuoteListener(ByVal pListener As IQuoteListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_Finish()
mMarketDataRequested = False
Set mMktDataManagerRef = Nothing
End Sub

Private Sub IMarketDataSource_NotifyCurrentDOM(ByVal pListener As IMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataQuoteValues(ByVal pListener As IQuoteListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDataTickValues(ByVal pListener As IGenericTickListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_NotifyCurrentMarketDepthValues(ByVal pListener As IRawMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_Refresh()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RefreshPriceChange(ByVal pListener As IPriceChangeListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RefreshQuotes(ByVal pListener As IQuoteListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RefreshTicks(ByVal pListener As IGenericTickListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveErrorListener(ByVal pListener As IErrorListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveFromConfig()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveGenericTickListener(ByVal pListener As IGenericTickListener)
mGenericTickListeners.Remove pListener
End Sub

Private Sub IMarketDataSource_RemoveMarketDepthListener(ByVal pListener As IMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveOptionComputationListener(ByVal pListener As IOptionComputationListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemovePriceChangeListener(ByVal pListener As IPriceChangeListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveQuoteListener(ByVal pListener As IQuoteListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveRawMarketDepthListener(ByVal pListener As IRawMarketDepthListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_StartMarketData()
mMarketDataRequested = True
doStartMarketData
End Sub

Private Sub IMarketDataSource_StartMarketDepth()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_StopMarketData()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

Private Sub IMarketDataSource_StopMarketDepth()
Assert False, ErrorCodes.ErrUnsupportedOperationException, "Not implemented"
End Sub

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
If Not ev.Future.IsAvailable Then Exit Sub

Dim lContract As IContract: Set lContract = ev.Future.Value
Set mContractSpec = lContract.Specifier
doStartMarketData
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub FireTick(ByRef pTick As GenericTick)
If Not mMarketDataRequested Then Exit Sub

Dim ev As GenericTickEventData
Set ev.Source = Me
ev.StreamId = mStreamId
ev.Tick = pTick

Static sInit As Boolean
Static sCurrentGenericTickListeners() As Object
Static sSomeCurrentGenericTickListeners As Boolean

If Not sInit Or Not mGenericTickListeners.Valid Then
    sInit = True
    sSomeCurrentGenericTickListeners = mGenericTickListeners.GetCurrentListeners(sCurrentGenericTickListeners)
End If

If sSomeCurrentGenericTickListeners Then
    Dim lListener As IGenericTickListener
    Dim i As Long
    For i = 0 To UBound(sCurrentGenericTickListeners)
        Set lListener = sCurrentGenericTickListeners(i)
        If pTick.TickType = -1 Then
            lListener.NoMoreTicks ev
        Else
            lListener.NotifyTick ev
        End If
    Next
End If
End Sub

Friend Sub Initialise( _
                ByVal pContractFuture As IFuture, _
                ByVal pMarketDataManager As MockMarketDataManager, _
                ByVal pStreamId As Long)
Set mContractFuture = pContractFuture
Set mFutureWaiter = New FutureWaiter
mFutureWaiter.Add pContractFuture
Set mMktDataManagerRef = CreateWeakReference(pMarketDataManager)
mStreamId = pStreamId
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doStartMarketData()
If Not mMarketDataRequested Then Exit Sub
If mContractSpec Is Nothing Then Exit Sub

getMktDataManagerRef.PlayTicks mContractSpec, Me
End Sub

Private Function getMktDataManagerRef() As MockMarketDataManager
Set getMktDataManagerRef = mMktDataManagerRef.Target
End Function


