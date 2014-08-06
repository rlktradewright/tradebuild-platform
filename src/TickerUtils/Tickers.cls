VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements Enumerable
Implements IMarketDataManager

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged(ev As CollectionChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "Tickers"

Private Const ConfigSectionTicker           As String = "Ticker"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTickers                 As EnumerableCollection
Attribute mTickers.VB_VarHelpID = -1

Private mOrderManager                       As OrderManager
Private mMarketDataManager                  As IMarketDataManager
Private mHistDataStore                      As IHistoricalDataStore

Private mStudyLibraryManager                As StudyLibraryManager

Private mOrderSubmitterFactoryLive          As IOrderSubmitterFactory
Private mOrderSubmitterFactorySimulated     As IOrderSubmitterFactory
Private mName                               As String
Private mScopeName                          As String
Private mMoneyManager                       As IMoneyManager

Private mConfig                             As ConfigurationSection

Private mCollectionChangeListeners          As New Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mTickers = New EnumerableCollection
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Set Enumerable_Enumerator = GetEnumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean

End Function

Private Sub Enumerable_Remove(pCookie As Variant)

End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean

End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' IMarketDataManager Interface Members
'@================================================================================

Private Function IMarketDataManager_CreateMarketDataSource( _
                ByVal pContractFuture As IFuture, _
                ByVal pIsRecoverable As Boolean, _
                Optional ByVal pkey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As StateChangeListener) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_CreateMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_CreateMarketDataSource = CreateTicker(pContractFuture, pIsRecoverable, pkey, pUseLocalTimeForStudies, pStateChangeListener)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IMarketDataManager_DataSources() As EnumerableCollection
Set IMarketDataManager_DataSources = mTickers
End Property

Private Sub IMarketDataManager_Finish()
Const ProcName As String = "IMarketDataManager_Finish"
On Error GoTo Err

Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IMarketDataManager_GetMarketDataSource(ByVal pkey As String) As IMarketDataSource
Const ProcName As String = "IMarketDataManager_GetMarketDataSource"
On Error GoTo Err

Set IMarketDataManager_GetMarketDataSource = GetTicker(pkey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IMarketDataManager_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IMarketDataManager_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataManager_RemoveFromConfig()
Const ProcName As String = "IMarketDataManager_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mTickers Event Handlers
'@================================================================================

Private Sub mTickers_CollectionChanged(ev As CollectionChangeEventData)
Const ProcName As String = "mTickers_CollectionChanged"
On Error GoTo Err

fireCollectionChange ev.changeType, ev.AffectedItem

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enumerator() As Enumerator
Set Enumerator = GetEnumerator
End Property

Friend Property Get MoneyManager() As IMoneyManager
Set MoneyManager = mMoneyManager
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Get OrderManager() As OrderManager
Set OrderManager = mOrderManager
End Property

Friend Property Get OrderSubmitterLive() As IOrderSubmitter
Set OrderSubmitterLive = mOrderSubmitterFactoryLive.CreateOrderSubmitter
End Property

Friend Property Get OrderSubmitterSimulated() As IOrderSubmitter
Set OrderSubmitterSimulated = mOrderSubmitterFactorySimulated.CreateOrderSubmitter
End Property

Public Property Get ScopeName() As String
ScopeName = mScopeName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddCollectionChangeListener(ByVal pListener As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function CreateTicker( _
                ByVal pContractFuture As IFuture, _
                Optional ByVal pIsRecoverable As Boolean = True, _
                Optional ByVal pkey As String, _
                Optional ByVal pUseLocalTimeForStudies As Boolean, _
                Optional ByVal pStateChangeListener As StateChangeListener) As Ticker
Const ProcName As String = "CreateTicker"
On Error GoTo Err

If pkey = "" Then pkey = GenerateGUIDString

Dim lDataSource As IMarketDataSource
Set lDataSource = mMarketDataManager.CreateMarketDataSource(pContractFuture, pIsRecoverable, pkey, pUseLocalTimeForStudies, pStateChangeListener)

Set CreateTicker = New Ticker
CreateTicker.Initialise lDataSource, mHistDataStore, Me, Nothing
                
If pIsRecoverable And Not mConfig Is Nothing Then addTickerToConfig pkey

mTickers.Add CreateTicker, pkey

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Dim lCurrTickers As New EnumerableCollection

Dim lTicker As Ticker
For Each lTicker In mTickers
    lCurrTickers.Add lTicker
Next

For Each lTicker In lCurrTickers
    lTicker.FinishAndRemoveFromConfig False
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetEnumerator() As Enumerator
Attribute GetEnumerator.VB_UserMemId = -4
Attribute GetEnumerator.VB_MemberFlags = "40"
Set GetEnumerator = mTickers.Enumerator
End Function

Public Function GetTicker(ByVal pkey As String) As Ticker
Const ProcName As String = "GetTicker"
On Error GoTo Err

Set GetTicker = mTickers.Item(pkey)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pMarketDataManager As IMarketDataManager, _
                ByVal pStudyLibraryManager As StudyLibraryManager, _
                ByVal pHistDataStore As IHistoricalDataStore, _
                ByVal pOrderManager As OrderManager, _
                ByVal pOrderSubmitterLive As IOrderSubmitter, _
                ByVal pOrderSubmitterSimulated As IOrderSubmitter, _
                ByVal pName As String, _
                ByVal pScopeName As String, _
                ByVal pMoneyManager As IMoneyManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mMarketDataManager = pMarketDataManager
Set mStudyLibraryManager = pStudyLibraryManager
Set mHistDataStore = pHistDataStore

Set mOrderManager = pOrderManager
Set mOrderSubmitterLive = pOrderSubmitterLive
Set mOrderSubmitterSimulated = pOrderSubmitterSimulated

mName = pName
mScopeName = pScopeName
Set mMoneyManager = pMoneyManager

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

gLogger.Log "Loading tickers from config", ProcName, ModuleName

Set mConfig = pConfig

Dim lCs As ConfigurationSection
For Each lCs In mConfig
    Dim lKey As String
    lKey = lCs.InstanceQualifier
    
    gLogger.Log "Loading ticker: " & lKey, ProcName, ModuleName
    
    Dim lPositionManager As PositionManager
    If mOrderManager.PositionManagersLive.Contains(lKey) Then
        Set lPositionManager = mOrderManager.PositionManagersLive.Item(lKey)
        If lPositionManager.IsFinished Then Set lPositionManager = Nothing
    End If
    
    Dim lDataSource As IMarketDataSource
    Set lDataSource = mMarketDataManager.GetMarketDataSource(lKey)
    
    Dim lTicker As Ticker
    Set lTicker = New Ticker
    lTicker.Initialise lDataSource, mHistDataStore, Me, lPositionManager
                
    mTickers.Add lTicker, lKey
Next

gLogger.Log "Loaded tickers from config", ProcName, ModuleName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function RefreshPriceChange( _
                ByVal listener As IPriceChangeListener) As TaskController
Const ProcName As String = "RefreshPriceChange"
On Error GoTo Err

Dim pcrt As New PriceChangeRefreshTask

If mTickers.Count = 0 Then Exit Function
pcrt.Initialise mTickers, listener
Set RefreshPriceChange = StartTask(pcrt, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Public Function RefreshQuotes( _
                ByVal listener As IQuoteListener) As TaskController
Const ProcName As String = "RefreshQuotes"
On Error GoTo Err

Dim qrt As New QuotesRefreshTask

If mTickers.Count = 0 Then Exit Function
qrt.Initialise mTickers, listener
Set RefreshQuotes = StartTask(qrt, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
                
Public Sub Remove(ByVal pTicker As Ticker)
Const ProcName As String = "Remove"
On Error GoTo Err

pTicker.FinishAndRemoveFromConfig True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveCollectionChangeListener(ByVal pListener As CollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mCollectionChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveNoFinish(ByVal pTicker As Ticker)
Const ProcName As String = "RemoveNoFinish"
On Error GoTo Err

mTickers.Remove pTicker.Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveTickerFromConfig(ByVal pTicker As Ticker)
Const ProcName As String = "RemoveTickerFromConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Sub

mConfig.RemoveConfigurationSection ConfigSectionTicker & "(" & pTicker.Key & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function StartTickers( _
                ByVal pContractsFuture As IFuture, _
                Optional ByVal pStartMarketData As Boolean = True, _
                Optional ByVal pStartMarketDepth As Boolean = False, _
                Optional ByVal pInterval As Long = 50) As TaskController
Const ProcName As String = "StartTickers"
On Error GoTo Err

Dim lTickersStarter As New TickersStarterTask

lTickersStarter.Initialise pContractsFuture, _
                            Me, _
                            pStartMarketData, _
                            pStartMarketDepth, _
                            pInterval
Set StartTickers = StartTask(lTickersStarter, PriorityNormal)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addTickerToConfig( _
                ByVal pTickerKey As String) As ConfigurationSection
Const ProcName As String = "addTickerToConfig"
On Error GoTo Err

If mConfig Is Nothing Then Exit Function

Dim lCs As ConfigurationSection
Set lCs = mConfig.AddConfigurationSection(ConfigSectionTicker & "(" & pTickerKey & ")")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireCollectionChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As Ticker)
Const ProcName As String = "fireCollectionChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.changeType = changeType
Set ev.AffectedItem = affectedObject

Dim listener As CollectionChangeListener
For Each listener In mCollectionChangeListeners.CurrentListeners
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




