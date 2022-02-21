VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OptionParametersStore"
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

Public Type IndexEntry
'    StartIndex          As Long
'    EndIndex            As Long
'    NumberOfEntries     As Long
    ExpiriesBuilder     As ExpiriesBuilder
    StrikesBuilder      As StrikesBuilder
End Type

'Private Type ParameterEntry
'    Next                As Long
'    Expiry              As String
'    Strike              As Double
'End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "OptionParametersStore"

'@================================================================================
' Member variables
'@================================================================================

Private mExpiriesBuilder                            As New ExpiriesBuilder
Private mStrikesBuilder                             As New StrikesBuilder

Private mExpiriesIndex                              As SortedDictionary
Private mStrikesIndex                               As SortedDictionary

'Private mParameters()                               As ParameterEntry
'Private mParametersNextIndex                        As Long

Private mIsComplete                                   As Boolean

Private mCacheDate                                  As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mExpiriesIndex = CreateSortedDictionary(KeyTypeString)
Set mStrikesIndex = CreateSortedDictionary(KeyTypeFloat)
'ReDim mParameters(127) As ParameterEntry
mCacheDate = Now
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

Friend Property Get IsComplete() As Boolean
IsComplete = mIsComplete
End Property

Friend Property Get IsExpired() As Boolean
IsExpired = gIsCacheEntryExpired(mCacheDate)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddExpiry( _
                ByVal pExpiry As String)
Const ProcName As String = "AddExpiry"
On Error GoTo Err

Assert Not mIsComplete, "Parameters store already complete"

mExpiriesBuilder.Add pExpiry

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddParameter( _
                ByVal pStrike As Double, _
                ByVal pExpiry As String)
Const ProcName As String = "AddParameter"
On Error GoTo Err

Assert Not mIsComplete, "Parameters store already complete"

Dim lStrikeIndexEntry As IndexEntry
getIndexEntry mStrikesIndex, pStrike, lStrikeIndexEntry, True
'addEntryToList lStrikeIndexEntry, pStrike, pExpiry
lStrikeIndexEntry.ExpiriesBuilder.Add pExpiry

Dim lExpiryIndexEntry As IndexEntry
getIndexEntry mExpiriesIndex, pExpiry, lExpiryIndexEntry, False
'addEntryToList lStrikeIndexEntry, pStrike, pExpiry
lExpiryIndexEntry.StrikesBuilder.Add pStrike

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddStrike( _
                ByVal pStrike As Double)
Const ProcName As String = "AddStrike"
On Error GoTo Err

Assert Not mIsComplete, "Parameters store already complete"

mStrikesBuilder.Add pStrike

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Complete()
mIsComplete = True
End Sub

Friend Function GetExpiries() As Expiries
Const ProcName As String = "GetExpiries"
On Error GoTo Err

Assert mIsComplete, "Parameters store not complete"

Set GetExpiries = mExpiriesBuilder.Expiries

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetExpiriesForStrike(ByVal pStrike As Double) As Expiries
Const ProcName As String = "GetExpiriesForStrike"
On Error GoTo Err

Assert mIsComplete, "Parameters store not complete"

Dim lStrikeIndexEntry As IndexEntry
getIndexEntry mStrikesIndex, pStrike, lStrikeIndexEntry, True
Set GetExpiriesForStrike = lStrikeIndexEntry.ExpiriesBuilder.Expiries

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetStrikes() As Strikes
Const ProcName As String = "GetStrikes"
On Error GoTo Err

Assert mIsComplete, "Parameters store not complete"

Set GetStrikes = mStrikesBuilder.Strikes


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetStrikesForExpiry(ByVal pExpiry As String) As Enumerator
Const ProcName As String = "GetStrikesForExpiry"
On Error GoTo Err

Assert mIsComplete, "Parameters store not complete"

Dim lIndexEntry As IndexEntry
getIndexEntry mExpiriesIndex, pExpiry, lIndexEntry, False
Set GetStrikesForExpiry = lIndexEntry.StrikesBuilder.Strikes

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

'Private Sub addEntryToList( _
'                ByRef pIndexEntry As IndexEntry, _
'                ByVal pStrike As Double, _
'                ByVal pExpiry As String)
'Const ProcName As String = "addEntryToList"
'On Error GoTo Err
'
'Dim lNewEntryIndex As Long: lNewEntryIndex = allocateEntry
'mParameters(pIndexEntry.EndIndex).Next = lNewEntryIndex
'pIndexEntry.EndIndex = lNewEntryIndex
'
'mParameters(lNewEntryIndex).Expiry = pExpiry
'mParameters(lNewEntryIndex).Strike = pStrike
'
'pIndexEntry.NumberOfEntries = pIndexEntry.NumberOfEntries + 1
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'Private Function allocateEntry() As Long
'If mParametersNextIndex > UBound(mParameters) Then
'    ReDim Preserve mParameters(2 * (UBound(mParameters) + 1) - 1) As ParameterEntry
'End If
'allocateEntry = mParametersNextIndex
'mParametersNextIndex = mParametersNextIndex + 1
'End Function

Private Sub getIndexEntry( _
                ByRef pIndexDict As SortedDictionary, _
                ByVal pKey As Variant, _
                ByRef pHeader As IndexEntry, _
                ByVal pForExpiries As Boolean)
If Not pIndexDict.TryItem(pKey, pHeader) Then
    InitialiseHeader pHeader, pForExpiries
    pIndexDict.Add pHeader, pKey
End If
End Sub

Private Sub InitialiseHeader( _
                ByRef pHeader As IndexEntry, _
                ByVal pForExpiries As Boolean)
Const ProcName As String = "InitialiseHeader"
On Error GoTo Err

'pHeader.StartIndex = allocateEntry
'mParameters(pHeader.StartIndex).Next = NullIndex

'pHeader.EndIndex = pHeader.StartIndex

If pForExpiries Then
    Set pHeader.ExpiriesBuilder = New ExpiriesBuilder
Else
    Set pHeader.StrikesBuilder = New StrikesBuilder
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



