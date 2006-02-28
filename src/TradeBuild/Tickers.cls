VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 2  'vbComplexBound
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tickers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================


Event ContractAmbiguous( _
                ByVal pTicker As Ticker, _
                ByVal contractSpec As contractSpecifier)
                
Event ContractInvalid( _
                ByVal pTicker As Ticker, _
                ByVal contractSpec As contractSpecifier)
                

Event MarketDepthNotAvailable( _
                ByVal pTicker As Ticker, _
                ByVal reason As String)


Event TickerAdded( _
                ByVal pTicker As Ticker)

Event TickerReady( _
                ByVal pTicker As Ticker)

Event TickerRemoved( _
                ByVal pTicker As Ticker)



'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================


' Be sure to keep this enum in step with TickerColumnsIDs
Private Enum TickerColumnNameIndexes
    
    columnName
    ColumnDescription
    ColumnSymbol
    ColumnSecType
    ColumnExpiry
    ColumnExchange
    ColumnCurrency
    ColumnRight
    ColumnStrike
    
    
    ColumnBid
    ColumnBidSize
    ColumnAsk
    ColumnAskSize
    ColumnTrade
    ColumnTradeSize
    ColumnVolume
    ColumnHigh
    ColumnLow
    ColumnClose
    ColumnChange
    ColumnChangePercent
    ColumnOpenInterest
    
    ColumnNameIndexMax
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTickers As Collection
Private WithEvents mRS As Recordset
Attribute mRS.VB_VarHelpID = -1
'Private mTickerTable() As TickerTableEntry
Private mColumnNames(TickerColumnNameIndexes.ColumnNameIndexMax) As String
Private mColumnMask As Long     ' additive values from TickerColumnIds enum
Private mColumnIDs(TickerColumnNameIndexes.ColumnNameIndexMax) As TickerColumnIds
                                ' records the order in which columns were added
Private mMaxColumnIDIndex As Long

Private mDeleting As Boolean    ' indicates that a recordset entry is being deleted


'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_GetDataMember(DataMember As String, data As Object)
If mRS.State = adStateClosed Then
    err.Raise ErrorCodes.DatasourceHasNotBeenGenerated, _
            "TradeBuild.Tickers::GetDataMember", _
            "Datasource has not been generated"
End If
    
Set data = mRS
End Sub

Private Sub Class_Initialize()
'ReDim mTickerTable(InitialMaxTickers) As TickerTableEntry
Set mTickers = New Collection
mMaxColumnIDIndex = -1

Set mRS = New ADODB.Recordset

' Set the default columns of the Recordset.
AddColumn TickerColumnIds.columnName, "Name"
AddColumn TickerColumnIds.ColumnBidSize, "Bid size"
AddColumn TickerColumnIds.ColumnBid, "Bid"
AddColumn TickerColumnIds.ColumnAsk, "Ask"
AddColumn TickerColumnIds.ColumnAskSize, "Ask size"
AddColumn TickerColumnIds.ColumnTrade, "Last"
AddColumn TickerColumnIds.ColumnTradeSize, "Last size"
AddColumn TickerColumnIds.ColumnVolume, "Volume"
AddColumn TickerColumnIds.ColumnHigh, "High"
AddColumn TickerColumnIds.ColumnLow, "Low"
AddColumn TickerColumnIds.ColumnClose, "Close"
AddColumn TickerColumnIds.ColumnSymbol, "Symbol"
AddColumn TickerColumnIds.ColumnSecType, "Sec type"
AddColumn TickerColumnIds.ColumnExpiry, "Expiry"
AddColumn TickerColumnIds.ColumnExchange, "Exchange"
AddColumn TickerColumnIds.ColumnCurrency, "Currency"
AddColumn TickerColumnIds.ColumnRight, "Right"
AddColumn TickerColumnIds.ColumnStrike, "Strike"

Generate

'mDefaultOutputTickfileFormat = TickfileTradeBuild
'mDefaultOutputTickFilePath = App.Path & "\"


End Sub

'================================================================================
' xxxx Interface Members
'================================================================================

'================================================================================
' mRS Event Handlers
'================================================================================

Private Sub mRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, _
                                ByVal cRecords As Long, _
                                adStatus As ADODB.EventStatusEnum, _
                                ByVal pRecordset As ADODB.Recordset)
Dim Ticker As Ticker

Select Case adReason
Case adRsnAddNew
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnDelete
    mDeleting = True
    Set Ticker = mTickers(mRS("Key"))
    If Ticker.State <> TickerStateCodes.Closing Then Ticker.StopTicker
    mDeleting = False
Case adRsnUpdate
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnUndoUpdate
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnUndoAddNew
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnUndoDelete
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnRequery
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnResynch
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnClose
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnMove
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnFirstChange
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnMoveFirst
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnMoveNext
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnMovePrevious
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
Case adRsnMoveLast
    adStatus = adStatusUnwantedEvent 'prevent subsequent notifications
End Select
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get Count() As Long
Count = mTickers.Count
End Property

Public Property Get Item(ByVal index As Variant) As Ticker
Attribute Item.VB_UserMemId = 0
Set Item = mTickers.Item(index)
End Property

'Friend Property Let MaxTickers(ByVal value As Long)
'ReDim Preserve mTickerTable(value - 1) As TickerTableEntry
'End Property

'================================================================================
' Methods
'================================================================================

Public Function add(Optional ByVal key As String, _
                    Optional ByVal position As Long = -1) As Ticker
Set add = New Ticker
add.Tickers = Me
If key = "" Then key = gGenerateGUIDString
add.key = key
add.position = position
If Not add Is Nothing Then
    mTickers.add add, key
    RaiseEvent TickerAdded(add)
End If

End Function

Public Sub AddColumn(ByVal columnId As TickerColumnIds, columnName As String)
Dim columnNameIndex As Long

If mRS.State <> adStateClosed Then
    err.Raise ErrorCodes.CantAddColumn, _
            "TradeBuild.Tickers::AddColumn", _
            "Can't add column in this context - call ClearColumns first"
End If

If (mColumnMask And columnId) <> 0 Then
    err.Raise ErrorCodes.ColumnAlreadyAdded, _
            "TradeBuild.Tickers::AddColumn", _
            "A column with this id has already been added"
End If

columnNameIndex = getColumnNameIndex(columnId)

If mColumnNames(columnNameIndex) <> "" Then
    err.Raise ErrorCodes.ColumnNameNotUnique, _
            "TradeBuild.Tickers::AddColumn", _
            "The column name is not unique"
End If

mMaxColumnIDIndex = mMaxColumnIDIndex + 1
mColumnIDs(mMaxColumnIDIndex) = columnId
mColumnNames(columnNameIndex) = columnName
mColumnMask = mColumnMask Or columnId
End Sub

Friend Sub ask(ByVal bookmark As Variant, _
                ByVal price As String, _
                ByVal size As Long)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnAsk) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnAsk)) = price
    End If
    If (mColumnMask And TickerColumnIds.ColumnAskSize) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnAskSize)) = size
    End If
    mRS.Update
End If
End Sub

Friend Sub bid(ByVal bookmark As Variant, _
                ByVal price As String, _
                ByVal size As Long)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnBid) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnBid)) = price
    End If
    If (mColumnMask And TickerColumnIds.ColumnBidSize <> 0) Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnBidSize)) = size
    End If
    mRS.Update
End If
End Sub

Public Sub ClearColumns()
Dim i As Long
mRS.Close
mColumnMask = 0
mMaxColumnIDIndex = -1
For i = 0 To UBound(mColumnNames)
    mColumnNames(i) = ""
    mColumnIDs(i) = 0
Next
End Sub

Friend Sub ContractAmbiguous(ByVal Ticker As Ticker, _
                            ByVal contractSpec As contractSpecifier)
RaiseEvent ContractAmbiguous(Ticker, contractSpec)
Hide Ticker
End Sub

Friend Sub ContractInvalid(ByVal Ticker As Ticker, _
                            ByVal contractSpec As contractSpecifier)
RaiseEvent ContractInvalid(Ticker, contractSpec)
Hide Ticker
End Sub

Public Function createTickFileManager() As TickFileManager
Set createTickFileManager = New TickFileManager
createTickFileManager.Tickers = Me
End Function

Public Sub Generate()
Dim i As Long
Dim columnNameIndex As TickerColumnNameIndexes
Dim Ticker As Ticker

If mRS.State <> adStateClosed Then
    err.Raise ErrorCodes.CantGenerateColumns, _
            "TradeBuild.Tickers::Generate", _
            "Can't generate columns - already generated"
End If

' Set the properties of the Recordset.
With mRS.fields
    .Append "Key", adVarChar, 255
    .Append "Position", adInteger
    
    For i = 0 To mMaxColumnIDIndex
        
        columnNameIndex = getColumnNameIndex(mColumnIDs(i))
        
        Select Case columnNameIndex
        Case columnName
            .Append mColumnNames(columnNameIndex), adVarChar, 255
        Case ColumnDescription
            .Append mColumnNames(columnNameIndex), adVarChar, 255
        Case ColumnSymbol
            .Append mColumnNames(columnNameIndex), adVarChar, 6
        Case ColumnSecType
            .Append mColumnNames(columnNameIndex), adVarChar, 5
        Case ColumnExpiry
            .Append mColumnNames(columnNameIndex), adVarChar, 8
        Case ColumnExchange
            .Append mColumnNames(columnNameIndex), adVarChar, 10
        Case ColumnCurrency
            .Append mColumnNames(columnNameIndex), adVarChar, 3
        Case ColumnRight
            .Append mColumnNames(columnNameIndex), adVarChar, 4
        Case ColumnStrike
            .Append mColumnNames(columnNameIndex), adDouble
        Case ColumnBid
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnBidSize
            .Append mColumnNames(columnNameIndex), adInteger
        Case ColumnAsk
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnAskSize
            .Append mColumnNames(columnNameIndex), adInteger
        Case ColumnTrade
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnTradeSize
            .Append mColumnNames(columnNameIndex), adInteger
        Case ColumnVolume
            .Append mColumnNames(columnNameIndex), adInteger
        Case ColumnHigh
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnLow
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnClose
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnChange
            .Append mColumnNames(columnNameIndex), adVarChar, 9
        Case ColumnChangePercent
            .Append mColumnNames(columnNameIndex), adVarChar, 7
        End Select
        
        DoEvents
        
    Next
End With

With mRS
    .CursorType = adOpenStatic
    .LockType = adLockOptimistic
    .Open
End With

For Each Ticker In mTickers
    setTickerFields Ticker
    DoEvents
Next
End Sub

Friend Sub Hide(ByVal pTicker As Ticker)
If mDeleting Then Exit Sub

On Error Resume Next            ' in case this ticker has already been hidden

mRS.bookmark = pTicker.bookmark
mRS.Delete
mRS.Update

End Sub

Friend Sub high(ByVal bookmark As Variant, _
                ByVal price As String)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnHigh) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnHigh)) = price
    End If
    mRS.Update
End If
End Sub

Friend Sub Last(ByVal bookmark As Variant, _
                ByVal price As String, _
                ByVal size As Long, _
                ByVal Change As String, _
                ByVal ChangePercent As Double)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnTrade) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnTrade)) = price
    End If
    If (mColumnMask And TickerColumnIds.ColumnTradeSize) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnTradeSize)) = size
    End If
    If (mColumnMask And TickerColumnIds.ColumnChange) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChange)) = Change
    End If
    If (mColumnMask And TickerColumnIds.ColumnChangePercent) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChangePercent)) = Format(ChangePercent, IIf(ChangePercent >= 0, "+", "-") & "0.00\%")
    End If
    mRS.Update
End If
End Sub

Friend Sub Low(ByVal bookmark As Variant, _
                ByVal price As String)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnLow) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnLow)) = price
    End If
    mRS.Update
End If
End Sub

Friend Sub MarketDepthNotAvailable( _
                ByVal pTicker As Ticker, _
                ByVal reason As String)
RaiseEvent MarketDepthNotAvailable(pTicker, reason)
End Sub

Public Sub MoveTicker(ByVal key As String, _
                    ByVal newPosition As Long)
Dim lTicker As Ticker

Set lTicker = mTickers(key)
lTicker.position = newPosition
mRS("Position") = newPosition
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mTickers.[_NewEnum]
End Function

Friend Sub openInterest(ByVal bookmark As Variant, _
                ByVal size As Long)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnOpenInterest) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnOpenInterest)) = size
    End If
    mRS.Update
End If
End Sub

Friend Sub PrevClose(ByVal bookmark As Variant, _
                ByVal price As String, _
                ByVal Change As String, _
                ByVal ChangePercent As Double)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnClose) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnClose)) = price
    End If
    If (mColumnMask And TickerColumnIds.ColumnChange) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChange)) = Change
    End If
    If (mColumnMask And TickerColumnIds.ColumnChangePercent) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChangePercent)) = Format(ChangePercent, IIf(ChangePercent >= 0, "+", "-") & "0.00\%")
    End If
    mRS.Update
End If
End Sub

Public Sub Remove(ByVal pTicker As Ticker)
If pTicker.State <> TickerStateCodes.Closing Then
    pTicker.StopTicker
    Exit Sub
End If
    
Hide pTicker
mTickers.Remove pTicker.key
pTicker.key = ""
RaiseEvent TickerRemoved(pTicker)
End Sub

Friend Sub StartTicker(ByVal pTicker As Ticker)

setTickerFields pTicker

RaiseEvent TickerReady(pTicker)
End Sub

Friend Sub volume(ByVal bookmark As Variant, _
                ByVal size As Long)
If mRS.State <> adStateClosed Then
    mRS.bookmark = bookmark
    If (mColumnMask And TickerColumnIds.ColumnVolume) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnVolume)) = size
    End If
    mRS.Update
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function getColumnNameIndex(ByVal columnId As TickerColumnIds) As TickerColumnNameIndexes
Select Case columnId
Case TickerColumnIds.columnName
    getColumnNameIndex = TickerColumnNameIndexes.columnName
Case TickerColumnIds.ColumnDescription
    getColumnNameIndex = TickerColumnNameIndexes.ColumnDescription
Case TickerColumnIds.ColumnSymbol
    getColumnNameIndex = TickerColumnNameIndexes.ColumnSymbol
Case TickerColumnIds.ColumnSecType
    getColumnNameIndex = TickerColumnNameIndexes.ColumnSecType
Case TickerColumnIds.ColumnExpiry
    getColumnNameIndex = TickerColumnNameIndexes.ColumnExpiry
Case TickerColumnIds.ColumnExchange
    getColumnNameIndex = TickerColumnNameIndexes.ColumnExchange
Case TickerColumnIds.ColumnCurrency
    getColumnNameIndex = TickerColumnNameIndexes.ColumnCurrency
Case TickerColumnIds.ColumnRight
    getColumnNameIndex = TickerColumnNameIndexes.ColumnRight
Case TickerColumnIds.ColumnStrike
    getColumnNameIndex = TickerColumnNameIndexes.ColumnStrike
Case TickerColumnIds.ColumnBid
    getColumnNameIndex = TickerColumnNameIndexes.ColumnBid
Case TickerColumnIds.ColumnBidSize
    getColumnNameIndex = TickerColumnNameIndexes.ColumnBidSize
Case TickerColumnIds.ColumnAsk
    getColumnNameIndex = TickerColumnNameIndexes.ColumnAsk
Case TickerColumnIds.ColumnAskSize
    getColumnNameIndex = TickerColumnNameIndexes.ColumnAskSize
Case TickerColumnIds.ColumnTrade
    getColumnNameIndex = TickerColumnNameIndexes.ColumnTrade
Case TickerColumnIds.ColumnTradeSize
    getColumnNameIndex = TickerColumnNameIndexes.ColumnTradeSize
Case TickerColumnIds.ColumnVolume
    getColumnNameIndex = TickerColumnNameIndexes.ColumnVolume
Case TickerColumnIds.ColumnHigh
    getColumnNameIndex = TickerColumnNameIndexes.ColumnHigh
Case TickerColumnIds.ColumnLow
    getColumnNameIndex = TickerColumnNameIndexes.ColumnLow
Case TickerColumnIds.ColumnClose
    getColumnNameIndex = TickerColumnNameIndexes.ColumnClose
Case TickerColumnIds.ColumnChange
    getColumnNameIndex = TickerColumnNameIndexes.ColumnChange
Case TickerColumnIds.ColumnChangePercent
    getColumnNameIndex = TickerColumnNameIndexes.ColumnChangePercent
End Select

End Function

Private Sub setTickerFields(ByVal pTicker As Ticker)
Dim priceFormatString As String
Dim bookmark As Variant

priceFormatString = pTicker.priceFormatString
bookmark = pTicker.bookmark
If IsEmpty(bookmark) Then
    mRS.AddNew
    pTicker.bookmark = mRS.bookmark
Else
    ' this happens when multiple tickfiles are being replayed on the
    ' second and subsequent tickfiles
    mRS.bookmark = bookmark
End If
mRS("Key") = pTicker.key
mRS("Position") = pTicker.position
With pTicker.contract
    
    If (mColumnMask And TickerColumnIds.columnName) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.columnName)) = .specifier.localSymbol
    End If
    If (mColumnMask And TickerColumnIds.ColumnDescription) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnDescription)) = .Description
    End If
    If (mColumnMask And TickerColumnIds.ColumnSymbol) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnSymbol)) = .specifier.symbol
    End If
    If (mColumnMask And TickerColumnIds.ColumnSecType) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnSecType)) = secTypeToString(.specifier.secType)
    End If
    If (mColumnMask And TickerColumnIds.ColumnExpiry) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnExpiry)) = .specifier.expiry
    End If
    If (mColumnMask And TickerColumnIds.ColumnExchange) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnExchange)) = .specifier.exchange
    End If
    If (mColumnMask And TickerColumnIds.ColumnCurrency) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnCurrency)) = .specifier.currencyCode
    End If
    If (mColumnMask And TickerColumnIds.ColumnRight) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnRight)) = OptRightToString(.specifier.right)
    End If
    If (mColumnMask And TickerColumnIds.ColumnStrike) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnStrike)) = .specifier.strike
    End If
End With

With pTicker
    If (mColumnMask And TickerColumnIds.ColumnBid) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnBid)) = Format(.BidPrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnBidSize) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnBidSize)) = .bidSize
    End If
    If (mColumnMask And TickerColumnIds.ColumnAsk) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnAsk)) = Format(.AskPrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnAskSize) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnAskSize)) = .AskSize
    End If
    If (mColumnMask And TickerColumnIds.ColumnTrade) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnTrade)) = Format(.TradePrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnTradeSize) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnTradeSize)) = .TradeSize
    End If
    If (mColumnMask And TickerColumnIds.ColumnVolume) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnVolume)) = .volume
    End If
    If (mColumnMask And TickerColumnIds.ColumnHigh) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnHigh)) = Format(.highPrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnLow) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnLow)) = Format(.lowPrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnClose) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnClose)) = Format(.closePrice, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnChange) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChange)) = Format(.Change, priceFormatString)
    End If
    If (mColumnMask And TickerColumnIds.ColumnChangePercent) <> 0 Then
        mRS(mColumnNames(TickerColumnNameIndexes.ColumnChangePercent)) = Format(.ChangePercent, "+0.00")
    End If
    
End With

mRS.Update

End Sub


