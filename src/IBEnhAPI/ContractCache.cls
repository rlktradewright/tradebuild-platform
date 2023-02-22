VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContractCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ContractCache"

Private Const ContinuationDataTypeContracts         As String = "CONTRACTS"
Private Const ContinuationDataTypeRuleset           As String = "RULESET"
Private Const ContinuationDataTypeOptionParams      As String = "OPTIONPARAMS"

'@================================================================================
' Member variables
'@================================================================================

Private mContracts                                  As New EnumerableCollection

Private mOptionParameters                           As New EnumerableCollection

Private mPriceIncrementRulesets                     As New EnumerableCollection

Private WithEvents mFutureWaiter                    As FutureWaiter
Attribute mFutureWaiter.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFutureWaiter = New FutureWaiter
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mFutureWaiter Event Handlers
'@================================================================================

Private Sub mFutureWaiter_WaitCompleted(ev As FutureWaitCompletedEventData)
Const ProcName As String = "mFutureWaiter_WaitCompleted"
On Error GoTo Err

Dim lContinuationData As CacheContinuationData
Set lContinuationData = ev.ContinuationData

Dim lKey As String

If lContinuationData.DataType = ContinuationDataTypeRuleset Then
    lKey = CStr(lContinuationData.DataValue)
    Dim lRulesetEntry As PriceIncrementRulesetCacheEntry
    lRulesetEntry = mPriceIncrementRulesets.Item(lKey)
    mPriceIncrementRulesets.Remove lKey
    
    If ev.Future.IsCancelled Or ev.Future.IsFaulted Then Exit Sub
    
    Set lRulesetEntry.RulesetFutureBuilder = Nothing
    Set lRulesetEntry.Ruleset = ev.Future.Value
    
    lRulesetEntry.CacheDate = Now
    
    mPriceIncrementRulesets.Add lRulesetEntry, lKey
    gLogger.Log "Added price increment ruleset " & _
                lKey & _
                " to cache", _
                ProcName, ModuleName, LogLevelDetail
ElseIf lContinuationData.DataType = ContinuationDataTypeContracts Then
    lKey = lContinuationData.DataValue
    Dim lContractEntry As ContractCacheEntry
    lContractEntry = mContracts.Item(lKey)
    mContracts.Remove lKey
    
    If ev.Future.IsCancelled Or ev.Future.IsFaulted Then Exit Sub
    
    Set lContractEntry.ContractsFutureBuilder = Nothing
    Set lContractEntry.Contracts = ev.Future.Value
    
    lContractEntry.CacheDate = Now
    
    mContracts.Add lContractEntry, lKey
    gLogger.Log "Added " & _
                lContractEntry.Contracts.Count & _
                IIf(lContractEntry.Contracts.Count = 1, " contract", " contracts") & _
                " to cache for key", _
                ProcName, ModuleName, LogLevelDetail, _
                lKey
ElseIf lContinuationData.DataType = ContinuationDataTypeOptionParams Then
    Dim lContractSpec As IContractSpecifier
    Set lContractSpec = lContinuationData.DataValue
    
    lKey = generateEndOptionParametersKey(lContractSpec)
    
    Dim lOptionParamsEntry As OptionParametersCacheEntry
    lOptionParamsEntry = mOptionParameters.Item(lKey)
    mOptionParameters.Remove lKey
    
    If ev.Future.IsCancelled Or ev.Future.IsFaulted Then Exit Sub
    
    Set lOptionParamsEntry.OptionParametersFutureBuilder = Nothing
    lOptionParamsEntry.CacheDate = Now
    
    mOptionParameters.Add lOptionParamsEntry, lKey
    
    gLogger.Log "Added option parameters to cache for contract", _
                ProcName, ModuleName, LogLevelDetail, _
                lContractSpec.Key
Else
    Assert False, "Invalid continuation data"
End If

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

Friend Sub AddContract( _
                ByVal pKey As String, _
                ByVal pContract As IContract, _
                ByVal pTwsContract As TwsContract)
Const ProcName As String = "AddContract"
On Error GoTo Err

If pContract.Specifier.SecType = SecTypeCombo Then Exit Sub
If Not IsValidExchangeCode(pTwsContract.Specifier.Exchange) Then Exit Sub

Dim lEntry As ContractCacheEntry
If tryGetContractCacheEntry(pKey, lEntry) Then Exit Sub

lEntry.CacheDate = Now
Dim lContractsBuilder As New ContractsBuilder
lContractsBuilder.Add pContract
lContractsBuilder.ContractSpecifier = pContract.Specifier
Set lEntry.Contracts = lContractsBuilder.Contracts
If lEntry.TwsContracts Is Nothing Then Set lEntry.TwsContracts = New EnumerableCollection
lEntry.TwsContracts.Add pTwsContract
mContracts.Add lEntry, pKey
gLogger.Log "Added 1 contract to cache for key: " & pKey, ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Friend Sub AddContracts( _
'                ByVal pContracts As IContracts, _
'                ByVal pTwsContracts As EnumerableCollection, _
'                ByVal pRequestContractSpec As IContractSpecifier)
'Const ProcName As String = "AddContracts"
'On Error GoTo Err
'
'Dim lEntry As ContractCacheEntry
'If tryGetContractCacheEntry(pRequestContractSpec.Key, lEntry) Then mContracts.Remove pRequestContractSpec.Key
'
'lEntry.CacheDate = Now
'Set lEntry.Contracts = pContracts
'Set lEntry.TwsContracts = pTwsContracts
'mContracts.Add lEntry, pRequestContractSpec.Key
'gLogger.Log "Added " & _
'            pContracts.Count & _
'            IIf(pContracts.Count = 1, " contract", " contracts") & _
'            " to cache for key", _
'            ProcName, ModuleName, LogLevelDetail, _
'            pRequestContractSpec.Key
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Friend Sub AddOptionParameter( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pExchange As String, _
                pExpirations() As String, _
                pStrikes() As Double)
Const ProcName As String = "AddOptionParameter"
On Error GoTo Err

Dim lEntry As OptionParametersCacheEntry
Dim lKey As String: lKey = generateOptionParametersKey(pContractSpec, pExchange)
If tryGetOptionParametersCacheEntry(lKey, lEntry) Then Exit Sub

lEntry.CacheDate = Now
Dim lParametersStore As New OptionParametersStore
Set lEntry.OptionParameters = lParametersStore

mOptionParameters.Add lEntry, lKey

Dim i As Long
Dim j As Long
For i = 0 To UBound(pExpirations)
    For j = 0 To UBound(pStrikes)
        lParametersStore.AddParameter pStrikes(j), pExpirations(i)
    Next
Next

For i = 0 To UBound(pExpirations)
    lParametersStore.AddExpiry pExpirations(i)
Next

For j = 0 To UBound(pStrikes)
    lParametersStore.AddStrike pStrikes(j)
Next

lParametersStore.Complete

gLogger.Log "Added option parameters to cache for key: " & lKey, ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddPriceIncrementRuleset( _
                ByVal pRuleset As PriceIncrementRuleset)
Const ProcName As String = "AddPriceIncrementRuleset"
On Error GoTo Err

Dim lEntry As PriceIncrementRulesetCacheEntry
If tryGetRulesetCacheEntry(CStr(pRuleset.RulesID), lEntry) Then Exit Sub

lEntry.CacheDate = Now
Set lEntry.Ruleset = pRuleset
mPriceIncrementRulesets.Add lEntry, CStr(pRuleset.RulesID)

gLogger.Log "Added price increment ruleset " & _
            pRuleset.RulesID & _
            " to cache", _
            ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ContainsContract( _
                ByVal pKey As String) As Boolean
Const ProcName As String = "ContainsContract"
On Error GoTo Err

Dim lEntry As ContractCacheEntry
ContainsContract = tryGetContractCacheEntry(pKey, lEntry)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetContract( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pContract As IContract, _
                ByRef pTwsContract As TwsContract) As Boolean
Const ProcName As String = "GetContract"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As ContractCacheEntry
If Not tryGetContractCacheEntry(pContractSpec.Key, lEntry) Then Exit Function

If Not lEntry.ContractsFutureBuilder Is Nothing Then
Else
    Assert lEntry.Contracts.Count = 1, "Contract is not uniquely specified"
    GetContract = True
    Set pContract = lEntry.Contracts.ItemAtIndex(1)
    Set pTwsContract = lEntry.TwsContracts(1)
    gLogger.Log "Retrieved 1 contract from cache for key", _
                ProcName, ModuleName, LogLevelMediumDetail, _
                pContractSpec.Key
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetContracts( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pContracts As IContracts, _
                ByRef pTwsContracts As EnumerableCollection) As Boolean
Const ProcName As String = "GetContracts"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As ContractCacheEntry
If Not tryGetContractCacheEntry(pContractSpec.Key, lEntry) Then Exit Function

If Not lEntry.ContractsFutureBuilder Is Nothing Then
Else
    GetContracts = True
    Set pContracts = lEntry.Contracts
    Set pTwsContracts = lEntry.TwsContracts
    gLogger.Log "Retrieved " & _
                lEntry.Contracts.Count & _
                IIf(lEntry.Contracts.Count = 1, " contract", " contracts") & _
                " from cache for key", _
                ProcName, ModuleName, LogLevelMediumDetail, _
                pContractSpec.Key
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetContractsFutureBuilder( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pContractsFutureBuilder As FutureBuilder, _
                ByRef pTwsContracts As EnumerableCollection) As Boolean
Const ProcName As String = "GetContractsFutureBuilder"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As ContractCacheEntry
If Not tryGetContractCacheEntry(pContractSpec.Key, lEntry) Then
    Set pContractsFutureBuilder = Nothing
    Set pTwsContracts = Nothing
    GetContractsFutureBuilder = False
Else
    Set pContractsFutureBuilder = lEntry.ContractsFutureBuilder
    Set pTwsContracts = lEntry.TwsContracts
    GetContractsFutureBuilder = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetOptionExpiries( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pStrike As Double, _
                ByRef pExpiries As Expiries) As Boolean
Const ProcName As String = "GetOptionExpiries"
On Error GoTo Err

Dim lOptionParametersStore As OptionParametersStore
Set lOptionParametersStore = getOptionParameters(pContractSpecifier, pExchange)

If lOptionParametersStore Is Nothing Then Exit Function

If pStrike = 0 Then
    Set pExpiries = lOptionParametersStore.GetExpiries()
Else
    Set pExpiries = lOptionParametersStore.GetExpiriesForStrike(pStrike)
End If

GetOptionExpiries = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetOptionStrikes( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String, _
                ByVal pExpiry As String, _
                ByRef pStrikes As Strikes) As Boolean
Const ProcName As String = "GetOptionStrikes"
On Error GoTo Err

Dim lOptionParametersStore As OptionParametersStore
Set lOptionParametersStore = getOptionParameters(pContractSpecifier, pExchange)

If lOptionParametersStore Is Nothing Then Exit Function

If pExpiry = "" Then
    Set pStrikes = lOptionParametersStore.GetStrikes()
Else
    Set pStrikes = lOptionParametersStore.GetStrikesForExpiry(pExpiry)
End If

GetOptionStrikes = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetOptionParametersFutureBuilder( _
                ByVal pContractSpec As IContractSpecifier, _
                ByRef pOptionParametersFutureBuilder As FutureBuilder) As Boolean
Const ProcName As String = "GetOptionParametersFutureBuilder"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Function

Dim lEntry As OptionParametersCacheEntry
If Not tryGetOptionParametersCacheEntry(generateEndOptionParametersKey(pContractSpec), lEntry) Then
    Set pOptionParametersFutureBuilder = Nothing
    GetOptionParametersFutureBuilder = False
Else
    Set pOptionParametersFutureBuilder = lEntry.OptionParametersFutureBuilder
    GetOptionParametersFutureBuilder = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetPriceIncrementRuleset( _
                ByVal pRuleID As Long) As PriceIncrementRuleset
Const ProcName As String = "GetPriceIncrementRuleset"
On Error GoTo Err

Dim lEntry As PriceIncrementRulesetCacheEntry
If Not tryGetRulesetCacheEntry(pRuleID, lEntry) Then Exit Function

If Not lEntry.Ruleset Is Nothing Then
    Set GetPriceIncrementRuleset = lEntry.Ruleset
    gLogger.Log "Retrieved price increment ruleset " & _
                pRuleID & _
                " from cache", _
                ProcName, ModuleName, LogLevelMediumDetail
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetRulesetFutureBuilder( _
                ByVal pRuleID As Long, _
                ByRef pRulesetFutureBuilder As FutureBuilder) As Boolean
Const ProcName As String = "GetRulesetFutureBuilder"
On Error GoTo Err

Dim lEntry As PriceIncrementRulesetCacheEntry
If Not tryGetRulesetCacheEntry(pRuleID, lEntry) Then
    Set pRulesetFutureBuilder = Nothing
    GetRulesetFutureBuilder = False
Else
    Set pRulesetFutureBuilder = lEntry.RulesetFutureBuilder
    GetRulesetFutureBuilder = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub SetContractsFutureBuilder( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pContractsFutureBuilder As FutureBuilder, _
                ByVal pTwsContracts As EnumerableCollection)
Const ProcName As String = "SetContractsFutureBuilder"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Sub

Dim lEntry As ContractCacheEntry
Assert Not tryGetContractCacheEntry(pContractSpec.Key, lEntry), "Cache entry already exists for: " & pContractSpec.Key
    
Set lEntry.ContractsFutureBuilder = pContractsFutureBuilder
Set lEntry.TwsContracts = pTwsContracts
lEntry.CacheDate = Now
mContracts.Add lEntry, pContractSpec.Key

Dim c As New CacheContinuationData
c.DataType = ContinuationDataTypeContracts
c.DataValue = pContractSpec.Key
mFutureWaiter.Add pContractsFutureBuilder.Future, c

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetOptionParametersFutureBuilder( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pOptionParametersFutureBuilder As FutureBuilder)
Const ProcName As String = "SetOptionParametersFutureBuilder"
On Error GoTo Err

If pContractSpec.SecType = SecTypeCombo Then Exit Sub

Dim lKey As String: lKey = generateEndOptionParametersKey(pContractSpec)

Dim lEntry As OptionParametersCacheEntry
Assert Not tryGetOptionParametersCacheEntry(lKey, lEntry), "OPtion parameters cache entry already exists for: " & pContractSpec.Key
    
Set lEntry.OptionParametersFutureBuilder = pOptionParametersFutureBuilder
lEntry.CacheDate = Now

mOptionParameters.Add lEntry, lKey

Dim c As New CacheContinuationData
c.DataType = ContinuationDataTypeOptionParams
Set c.DataValue = pContractSpec
mFutureWaiter.Add pOptionParametersFutureBuilder.Future, c

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetRulesetFutureBuilder( _
                ByVal pRuleID As Long, _
                ByVal pRulesetFutureBuilder As FutureBuilder)
Const ProcName As String = "SetRulesetFutureBuilder"
On Error GoTo Err

Dim lEntry As PriceIncrementRulesetCacheEntry
Assert Not tryGetRulesetCacheEntry(pRuleID, lEntry), "Cache entry already exists for ruleset: " & pRuleID
    
Set lEntry.RulesetFutureBuilder = pRulesetFutureBuilder
lEntry.CacheDate = Now
mPriceIncrementRulesets.Add lEntry, CStr(Abs(pRuleID))

Dim c As New CacheContinuationData
c.DataType = ContinuationDataTypeRuleset
c.DataValue = Abs(pRuleID)
mFutureWaiter.Add pRulesetFutureBuilder.Future, c

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateEndOptionParametersKey( _
                ByVal pContractSpec As IContractSpecifier)
generateEndOptionParametersKey = pContractSpec.Key & "|END"
End Function

Private Function generateOptionParametersKey( _
                ByVal pContractSpec As IContractSpecifier, _
                ByVal pExchange As String)
Const ProcName As String = "generateOptionParametersKey"
On Error GoTo Err

generateOptionParametersKey = pContractSpec.Key & "|" & pExchange

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getOptionParameters( _
                ByVal pContractSpecifier As IContractSpecifier, _
                ByVal pExchange As String) As OptionParametersStore
Const ProcName As String = "getOptionParameters"
On Error GoTo Err

Dim lEntry As OptionParametersCacheEntry
If Not tryGetOptionParametersCacheEntry( _
                generateOptionParametersKey(pContractSpecifier, pExchange), _
                lEntry) Then Exit Function

Set getOptionParameters = lEntry.OptionParameters

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tryGetContractCacheEntry( _
                ByVal pKey As String, _
                ByRef pCacheEntry As ContractCacheEntry) As Boolean
Const ProcName As String = "tryGetContractCacheEntry"
On Error GoTo Err

If Not mContracts.TryItem(pKey, pCacheEntry) Then
ElseIf gIsCacheEntryExpired(pCacheEntry.CacheDate) Then
    mContracts.Remove pKey
Else
    tryGetContractCacheEntry = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tryGetOptionParametersCacheEntry( _
                ByVal pKey As String, _
                ByRef pCacheEntry As OptionParametersCacheEntry) As Boolean
Const ProcName As String = "tryGetOptionParametersCacheEntry"
On Error GoTo Err

If Not mOptionParameters.TryItem(pKey, pCacheEntry) Then
ElseIf gIsCacheEntryExpired(pCacheEntry.CacheDate) Then
    mOptionParameters.Remove pKey
Else
    tryGetOptionParametersCacheEntry = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function tryGetRulesetCacheEntry( _
                ByVal pRuleID As Long, _
                ByRef pCacheEntry As PriceIncrementRulesetCacheEntry) As Boolean
Const ProcName As String = "tryGetRulesetCacheEntry"
On Error GoTo Err

If Not mPriceIncrementRulesets.TryItem(CStr(Abs(pRuleID)), pCacheEntry) Then

' don't check for ruleset expiry to avoid problems where a ticker is
' started before midnight but the first tick doesn't arrive till after
' midnight
'ElseIf gIsCacheEntryExpired(pCacheEntry.CacheDate) Then
'    Dim lCacheEntry As PriceIncrementRulesetCacheEntry
'    pCacheEntry = lCacheEntry
'    mPriceIncrementRulesets.Remove CStr(Abs(pRuleID))
Else
    tryGetRulesetCacheEntry = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


