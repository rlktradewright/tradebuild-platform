VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CurrencyConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Interfaces
'@================================================================================

Implements ICurrencyConverter
Implements IGenericTickListener
Implements IMarketDataErrorListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ConversionTableEntry
    State                   As CurrencyConverterStates
    InvalidConversion       As Boolean
    ConversionFactor        As Double
    ReverseEntryIndex       As Long
    ContractFuture          As IFuture
    ContractInvalid         As Boolean
    Bid                     As Double
    Ask                     As Double
    Listeners               As Listeners
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "CurrencyConverter"

'@================================================================================
' Member variables
'@================================================================================

Private mContractStore                              As IContractStore
Private mMarketDataFactory                          As MarketDataFactory

' maps currency pair to conversion table index
Private mIndexes                                    As New EnumerableCollection

' maps tickstream id to conversion table index
Private mStreamIndexes                              As New EnumerableCollection

' maps stream ids to conversion factors
Private mConversionTable()                          As ConversionTableEntry
Private mNextConversionTableIndex                   As Long

Private WithEvents mResubscribeTimer                As IntervalTimer
Attribute mResubscribeTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mConversionTable(7) As ConversionTableEntry
End Sub

'@================================================================================
' ICurrencyConverter Interface Members
'@================================================================================

Private Sub ICurrencyConverter_AddStateChangeListener( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String, _
                ByVal pListener As IStateChangeListener)
Const ProcName As String = "ICurrencyConverter_AddStateChangeListener"
On Error GoTo Err

Dim lKey As String: lKey = getCurrencyPairKey( _
                                            pFromCurrency, _
                                            pToCurrency)
Dim lIndex As Long
AssertArgument mIndexes.TryItem(lKey, lIndex), "Not ready for currency pair " & lKey

With mConversionTable(lIndex)
    If .Listeners Is Nothing Then Set .Listeners = New Listeners
    .Listeners.Add pListener
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ICurrencyConverter_Convert( _
                ByVal pAmount As Double, _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As Double
Const ProcName As String = "ICurrencyConverter_Convert"
On Error GoTo Err

Dim lKey As String: lKey = getCurrencyPairKey( _
                                            pFromCurrency, _
                                            pToCurrency)

Dim lIndex As Long
AssertArgument mIndexes.TryItem(lKey, lIndex), "Not ready for currency pair " & lKey

ICurrencyConverter_Convert = pAmount * mConversionTable(lIndex).ConversionFactor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ICurrencyConverter_IsReadyCurrencies( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As Boolean
Const ProcName As String = "ICurrencyConverter_IsReadyCurrency"
On Error GoTo Err

Dim lKey As String: lKey = getCurrencyPairKey( _
                                            pFromCurrency, _
                                            pToCurrency)

Dim lIndex As Long
If mIndexes.TryItem(lKey, lIndex) Then
    AssertArgument Not mConversionTable(lIndex).InvalidConversion, "Invalid currency conversion"
    ICurrencyConverter_IsReadyCurrencies = mConversionTable(lIndex).State
ElseIf startTickStreams(pFromCurrency, pToCurrency) Then
    ICurrencyConverter_IsReadyCurrencies = False
Else
    ' this is an 'identity' conversion
    ICurrencyConverter_IsReadyCurrencies = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ICurrencyConverter_RemoveStateChangeListener( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String, _
                ByVal pListener As IStateChangeListener)
Const ProcName As String = "ICurrencyConverter_RemoveStateChangeListener"
On Error GoTo Err

Dim lKey As String: lKey = getCurrencyPairKey( _
                                            pFromCurrency, _
                                            pToCurrency)
Dim lIndex As Long
If Not mIndexes.TryItem(lKey, lIndex) Then Exit Sub

With mConversionTable(lIndex)
    If Not .Listeners Is Nothing Then .Listeners.Remove pListener
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGenericTickListener Interface Members
'@================================================================================

Private Sub IGenericTickListener_NoMoreTicks(ev As GenericTickEventData)
mStreamIndexes.Remove CStr(ev.StreamId)
End Sub

Private Sub IGenericTickListener_NotifyTick(ev As GenericTickEventData)
Const ProcName As String = "IGenericTickListener_NotifyTick"
On Error GoTo Err

If mResubscribeTimer Is Nothing Then
    Set mResubscribeTimer = CreateIntervalTimer(1, ExpiryTimeUnitMinutes, 60000)
    mResubscribeTimer.StartTimer
End If
    
Dim lIndex As Long
lIndex = mStreamIndexes(CStr(ev.StreamId))

With mConversionTable(lIndex)
Dim lState As CurrencyConverterStates: lState = CurrencyConverterStateNotReady
    If ev.Tick.TickType = TickTypeBid Then
        .Bid = ev.Tick.Price
    ElseIf ev.Tick.TickType = TickTypeAsk Then
        .Ask = ev.Tick.Price
    Else
        Exit Sub
    End If
    
    If .Bid = 0 Or .Ask = 0 Then Exit Sub
            
    .ConversionFactor = (.Bid + .Ask) / 2#
    lState = CurrencyConverterStateReady
    
    If .ReverseEntryIndex <> NullIndex Then
        mConversionTable(.ReverseEntryIndex).InvalidConversion = False
        mConversionTable(.ReverseEntryIndex).ConversionFactor = 1# / .ConversionFactor
    End If
    
    If lState = CurrencyConverterStateReady Then
        setState lIndex, CurrencyConverterStateReady
        If .ReverseEntryIndex <> NullIndex Then setState .ReverseEntryIndex, CurrencyConverterStateReady
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ICurrencyConverter Interface Members
'@================================================================================

Private Sub IMarketDataErrorListener_NotifyMarketDataError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
Const ProcName As String = "IMarketDataErrorListener_NotifyMarketDataError"
On Error GoTo Err

If pErrorCode = 200 Then
    Dim lIndex As Long: lIndex = mStreamIndexes(CStr(pStreamId))
    mConversionTable(lIndex).ContractInvalid = True
    If mConversionTable(mConversionTable(lIndex).ReverseEntryIndex).ContractInvalid Then
        mConversionTable(mConversionTable(lIndex).ReverseEntryIndex).InvalidConversion = True
        mConversionTable(lIndex).InvalidConversion = True
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IMarketDataErrorListener_NotifyMarketDepthError(ByVal pStreamId As Long, ByVal pErrorCode As Long, ByVal pErrorMessage As String)
End Sub

'@================================================================================
' mResubscribeTimer Event Handlers
'@================================================================================

Private Sub mResubscribeTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mResubscribeTimer_TimerExpired"
On Error GoTo Err

Dim i As Long
For i = 0 To mNextConversionTableIndex - 1
    If mConversionTable(i).InvalidConversion Then
    ElseIf mConversionTable(i).ContractFuture Is Nothing Then
        ' this is an 'identity' conversion
    Else
        Dim lStream As ITickStream
        Set lStream = mMarketDataFactory.CreateSnapshotMarketDataStream( _
                                        mConversionTable(i).ContractFuture, _
                                        Me)
        lStream.AddGenericTickListener Me
        lStream.Start
        mStreamIndexes.Add i, CStr(lStream.StreamId)
    End If
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pContractStore As IContractStore, _
                ByVal pMarketDataFactory As MarketDataFactory)
Set mContractStore = pContractStore
Set mMarketDataFactory = pMarketDataFactory
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateIndex() As Long
If mNextConversionTableIndex > UBound(mConversionTable) Then
    ReDim Preserve mConversionTable(2 * (UBound(mConversionTable) + 1) - 1) As ConversionTableEntry
End If
allocateIndex = mNextConversionTableIndex
mNextConversionTableIndex = mNextConversionTableIndex + 1
End Function

Private Sub fireStateChange(ByVal pIndex As Long)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = mConversionTable(pIndex).State

If mConversionTable(pIndex).Listeners Is Nothing Then Exit Sub

Dim lListeners() As Object
If getStateChangeListeners(mConversionTable(pIndex).Listeners, lListeners) Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(lListeners)
        Set lListener = lListeners(i)
        lListener.Change ev
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getStateChangeListeners( _
                ByVal pListeners As Listeners, _
                ByRef pCurrentListeners() As Object) As Boolean
Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not pListeners.Valid Then
    sInit = True
    sSomeListeners = pListeners.GetCurrentListeners(sCurrentListeners)
End If

pCurrentListeners = sCurrentListeners
getStateChangeListeners = sSomeListeners
End Function

Private Function getCurrencyPairKey( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As String
getCurrencyPairKey = pFromCurrency & "." & pToCurrency
End Function

Private Sub setState( _
                ByVal pIndex As Long, _
                ByVal pState As CurrencyConverterStates)
With mConversionTable(pIndex)
    If pState = .State Then Exit Sub
    .State = pState
    fireStateChange pIndex
End With
End Sub

Private Function startTickStreams( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String) As Boolean
Const ProcName As String = "startTickStreams"
On Error GoTo Err

Dim lIndex As Long
lIndex = allocateIndex
mIndexes.Add lIndex, getCurrencyPairKey(pFromCurrency, pToCurrency)

If pFromCurrency = pToCurrency Then
    mConversionTable(lIndex).ConversionFactor = 1#
    mConversionTable(lIndex).ReverseEntryIndex = NullIndex
    mConversionTable(lIndex).State = CurrencyConverterStateReady
    startTickStreams = False
    Exit Function
End If

Dim lStream1 As ITickStream
Set lStream1 = startTickStream(pFromCurrency, pToCurrency, lIndex)
lStream1.AddGenericTickListener Me
mStreamIndexes.Add lIndex, CStr(lStream1.StreamId)

Dim lReverseIndex As Long
lReverseIndex = allocateIndex
mIndexes.Add lReverseIndex, getCurrencyPairKey(pToCurrency, pFromCurrency)

Dim lStream2 As ITickStream
Set lStream2 = startTickStream(pToCurrency, pFromCurrency, lReverseIndex)
lStream2.AddGenericTickListener Me
mStreamIndexes.Add lReverseIndex, CStr(lStream2.StreamId)

mConversionTable(lIndex).ReverseEntryIndex = lReverseIndex
mConversionTable(lReverseIndex).ReverseEntryIndex = lIndex

startTickStreams = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function startTickStream( _
                ByVal pFromCurrency As String, _
                ByVal pToCurrency As String, _
                ByVal pIndex As Long) As ITickStream
Const ProcName As String = "startTickStream"
On Error GoTo Err

Dim lContractSpec As IContractSpecifier
Set lContractSpec = CreateContractSpecifier( _
                            getCurrencyPairKey(pFromCurrency, pToCurrency), _
                            "", _
                            "IDEALPRO", _
                            SecurityTypes.SecTypeCash)

Dim lContractFuture As IFuture
Set lContractFuture = FetchContract(lContractSpec, mContractStore)
Set mConversionTable(pIndex).ContractFuture = lContractFuture

Set startTickStream = mMarketDataFactory.CreateSnapshotMarketDataStream( _
                                        lContractFuture, _
                                        Me)
startTickStream.Start

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

