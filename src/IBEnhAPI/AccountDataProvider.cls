VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AccountDataProvider"
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

Implements IAccountDataConsumer
Implements IAccountDataProvider
Implements ITwsConnectionStateListener

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

Private Const ModuleName                            As String = "AccountDataProvider"

'@================================================================================
' Member variables
'@================================================================================

Private mTwsAPI                                     As TwsAPI

Private mIsTwsConnectedToIBServers                  As Boolean
Private mState                                      As AccountProviderStates

Private mLoadRequested                              As Boolean
Private mKeepUpdated                                As Boolean

Private mAccountValues                              As SortedDictionary
Private mPortfolioValues                            As SortedDictionary

Private mAccountValueListeners                      As Listeners
Private mPortfolioValueListeners                    As Listeners
Private mStateChangeListeners                       As Listeners

Private mAccountNotReady                            As Boolean
Private mBaseCurrency                               As String

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = AccountProviderNotReady
Set mAccountValues = CreateSortedDictionary(KeyTypeString)
Set mPortfolioValues = CreateSortedDictionary(KeyTypeString)
Set mAccountValueListeners = New Listeners
Set mPortfolioValueListeners = New Listeners
Set mStateChangeListeners = New Listeners
End Sub

'@================================================================================
' IAccountDataConsumer Interface Members
'@================================================================================

Private Sub IAccountDataConsumer_EndAccountSummary(ByVal pRequestId As Long)

End Sub

Private Sub IAccountDataConsumer_EndAccountValue(ByVal pAccountName As String)
If mAccountNotReady Then Exit Sub

setState AccountProviderReady
fireAccountValueEnd
firePortfolioValueEnd
End Sub

Private Sub IAccountDataConsumer_EndPosition()

End Sub

Private Sub IAccountDataConsumer_NotifyAccountSummary( _
                ByVal pRequestId As Long, _
                ByVal pAccount As String, _
                ByVal pTag As String, _
                ByVal pValue As String, _
                ByVal pCurrency As String)

End Sub

Private Sub IAccountDataConsumer_NotifyAccountTime(ByVal pTimeStamp As String)

End Sub

Private Sub IAccountDataConsumer_NotifyAccountValue( _
                ByVal pKey As String, _
                ByVal pValue As String, _
                ByVal pCurrency As String, _
                ByVal pAccountName As String)
Const ProcName As String = "IAccountDataConsumer_NotifyAccountValue"
On Error GoTo Err

If pKey = "AccountReady" Then
    If pValue = "true" Then
        mAccountNotReady = False
    Else
        mAccountNotReady = True
        setState AccountProviderNotReady
    End If
    Exit Sub
End If

If pKey = "NetLiquidation" Then mBaseCurrency = pCurrency
If pKey = "EquityWithLoanValue" Then
    pKey = "EquityWithLoan"
End If
Dim lAccountValue As AccountValue
Dim lID As String: lID = getValueID(pKey, pCurrency)
If Not mAccountValues.TryItem(lID, lAccountValue) Then
    Set lAccountValue = New AccountValue
    mAccountValues.Add lAccountValue, lID
End If
lAccountValue.Initialise pAccountName, pCurrency, pKey, pValue
If Not mAccountNotReady Then fireAccountValue lAccountValue

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IAccountDataConsumer_NotifyPortfolioUpdate( _
                ByVal pTwsContractSpec As TwsContractSpecifier, _
                ByVal pPosition As BoxedDecimal, _
                ByVal pMarketPrice As Double, _
                ByVal pMarketValue As Double, _
                ByVal pAverageCost As Double, _
                ByVal pUnrealizedPNL As Double, _
                ByVal pRealizedPNL As Double, _
                ByVal pAccountName As String)
Const ProcName As String = "IAccountDataConsumer_NotifyPortfolioUpdate"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
Set lContractSpec = GIBEnhancedApi.TwsContractSpecToContractSpecifier(pTwsContractSpec, 1)

Dim lKey As String: lKey = getContractName(lContractSpec)
Dim lPortfolioValue As PortfolioValue
If Not mPortfolioValues.TryItem(lKey, lPortfolioValue) Then
    Set lPortfolioValue = New PortfolioValue
    mPortfolioValues.Add lPortfolioValue, lKey
End If
lPortfolioValue.Initialise pAccountName, pAverageCost, lContractSpec, pMarketPrice, pMarketValue, pPosition, pRealizedPNL, pUnrealizedPNL
firePortfolioValue lPortfolioValue

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IAccountDataConsumer_NotifyPosition(ByVal pAccount As String, ByVal pContractSpec As TwsContractSpecifier, ByVal pPosition As BoxedDecimal, ByVal pAverageCost As Double)

End Sub

'@================================================================================
' IAccountDataProvider Interface Members
'@================================================================================

Private Sub IAccountDataProvider_AddAccountValueListener(ByVal pListener As IAccountValueListener)
mAccountValueListeners.Add pListener
End Sub

Private Sub IAccountDataProvider_AddPortfolioValueListener(ByVal pListener As IPortfolioValueListener)
mPortfolioValueListeners.Add pListener
End Sub

Private Sub IAccountDataProvider_AddStateChangeListener(ByVal pListener As IStateChangeListener)
mStateChangeListeners.Add pListener
End Sub

Private Property Get IAccountDataProvider_BaseCurrency() As String
Const ProcName As String = "IAccountDataProvider_BaseCurrency"
On Error GoTo Err

Assert mState = AccountProviderReady, "Account provider is not ready"
IAccountDataProvider_BaseCurrency = mBaseCurrency

Exit Property

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Property

Private Function IAccountDataProvider_GetAccountValue(ByVal pKey As String, Optional ByVal pCurrency As String) As IAccountValue
Const ProcName As String = "IAccountDataProvider_GetAccountValue"
On Error GoTo Err

Assert mState = AccountProviderReady, "Account provider is not ready"
Set IAccountDataProvider_GetAccountValue = mAccountValues.Item(getValueID(pKey, pCurrency))

Exit Function

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Function

Private Function IAccountDataProvider_GetAccountValues() As Enumerator
Set IAccountDataProvider_GetAccountValues = mAccountValues.Enumerator
End Function

Private Function IAccountDataProvider_GetPortfolioValue(ByVal pContractName As String) As IPortfolioValue
Set IAccountDataProvider_GetPortfolioValue = mPortfolioValues.Item(pContractName)
End Function

Private Function IAccountDataProvider_GetPortfolioValues() As Enumerator
Set IAccountDataProvider_GetPortfolioValues = mPortfolioValues.Enumerator
End Function

Private Sub IAccountDataProvider_Load(ByVal pKeepUpdated As Boolean)
Const ProcName As String = "IAccountDataProvider_Load"
On Error GoTo Err

mKeepUpdated = pKeepUpdated
If mTwsAPI.IsTwsConnectedToIBServers Then
    mTwsAPI.RequestAccountUpdates mKeepUpdated, ""
Else
    mLoadRequested = True
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IAccountDataProvider_RemoveAccountValueListener(ByVal pListener As IAccountValueListener)
mAccountValueListeners.Remove pListener
End Sub

Private Sub IAccountDataProvider_RemovePortfolioValueListener(ByVal pListener As IPortfolioValueListener)
mPortfolioValueListeners.Remove pListener
End Sub

Private Sub IAccountDataProvider_RemoveStateChangeListener(ByVal pListener As TWUtilities40.IStateChangeListener)
mStateChangeListeners.Remove pListener
End Sub

Private Property Get IAccountDataProvider_State() As AccountProviderStates
IAccountDataProvider_State = mState
End Property

'@================================================================================
' ITwsConnectionStateListener Interface Members
'@================================================================================

Private Sub ITwsConnectionStateListener_NotifyAPIConnectionStateChange(ByVal pSource As Object, ByVal pState As ApiConnectionStates, ByVal pMessage As String)
Const ProcName As String = "ITwsConnectionStateListener_NotifyAPIConnectionStateChange"
On Error GoTo Err

If pState = TwsConnConnected Then
    ' we'll assume that TWS has a server connection: if not,
    ' it will soon tell us
    mIsTwsConnectedToIBServers = True
    
    If mLoadRequested Then mTwsAPI.RequestAccountUpdates mKeepUpdated, ""
ElseIf pState <> TwsConnConnected Then
    setState AccountProviderNotReady
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionClosed(ByVal pSource As Object)
mIsTwsConnectedToIBServers = False
setState AccountProviderNotReady
End Sub

Private Sub ITwsConnectionStateListener_NotifyIBServerConnectionRecovered(ByVal pSource As Object, ByVal pDataLost As Boolean)
Const ProcName As String = "ITwsConnectionStateListener_NotifyIBServerConnectionRecovered"
On Error GoTo Err

mIsTwsConnectedToIBServers = True
mTwsAPI.RequestAccountUpdates mKeepUpdated, ""

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pTwsAPI As TwsAPI)
Set mTwsAPI = pTwsAPI
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireAccountValue( _
                ByVal pAccountValue As AccountValue)
Const ProcName As String = "fireAccountValue"
On Error GoTo Err

Dim ev As AccountValueEventData
Set ev.Source = Me
Set ev.data = pAccountValue

Dim lListeners() As Object
If getAccountListeners(lListeners) Then
    Dim lListener As IAccountValueListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.NotifyAccountValue ev
    Next
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireAccountValueEnd()
Const ProcName As String = "fireAccountValueEnd"
On Error GoTo Err

Dim ev As AccountValueEventData
Set ev.Source = Me
Set ev.data = Nothing

Dim lListeners() As Object
If getAccountListeners(lListeners) Then
    Dim lListener As IAccountValueListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.EndAccountValue ev
    Next
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePortfolioValue( _
                ByVal pPortfolioValue As PortfolioValue)
Const ProcName As String = "firePortfolioValue"
On Error GoTo Err

Dim ev As PortfolioValueEventData
Set ev.Source = Me
Set ev.data = pPortfolioValue

Dim lListeners() As Object
If getPortfolioListeners(lListeners) Then
    Dim lListener As IPortfolioValueListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.NotifyPortfolioValue ev
    Next
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub firePortfolioValueEnd()
Const ProcName As String = "firePortfolioValueEnd"
On Error GoTo Err

Dim ev As PortfolioValueEventData
Set ev.Source = Me
Set ev.data = Nothing

Dim lListeners() As Object
If getPortfolioListeners(lListeners) Then
    Dim lListener As IPortfolioValueListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.EndPortfolioValue ev
    Next
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireStateChange()
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mState

Dim lListeners() As Object
If getStateChangeListeners(lListeners) Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.Change ev
    Next
End If

Exit Sub

Err:
GIBEnhApi.HandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getAccountListeners(ByRef pListeners() As Object) As Boolean
Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mAccountValueListeners.Valid Then
    sInit = True
    sSomeListeners = mAccountValueListeners.GetCurrentListeners(sCurrentListeners)
End If

pListeners = sCurrentListeners
getAccountListeners = sSomeListeners
End Function

Private Function getContractName(ByVal pContractSpec As IContractSpecifier) As String
getContractName = SecTypeToShortString(pContractSpec.SecType) & _
                    ":" & pContractSpec.LocalSymbol & _
                    "@" & pContractSpec.Exchange
End Function

Private Function getPortfolioListeners(ByRef pListeners() As Object) As Boolean
Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mPortfolioValueListeners.Valid Then
    sInit = True
    sSomeListeners = mPortfolioValueListeners.GetCurrentListeners(sCurrentListeners)
End If

pListeners = sCurrentListeners
getPortfolioListeners = sSomeListeners
End Function

Private Function getStateChangeListeners(ByRef pListeners() As Object) As Boolean
Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If

pListeners = sCurrentListeners
getStateChangeListeners = sSomeListeners
End Function

Private Function getValueID( _
                ByVal pKey As String, _
                ByVal pCurrency As String) As String
getValueID = UCase$(pKey & "||" & pCurrency)
End Function

Private Sub setState(ByVal pState As AccountProviderStates)
If pState = mState Then Exit Sub
mState = pState
fireStateChange
End Sub
