Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#

Public Const ContractInfoSPName As String = "IB TWS Contract Info Service Provider"
Public Const HistoricDataSPName As String = "IB TWS Historic Data Service Provider"

Public Const ProviderKey As String = "TWS"

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    ' generic run-time error codes defined by VB
    InvalidProcedureCall = 5
    Overflow = 6
    SubscriptOutOfRange = 9
    DivisionByZero = 11
    TypeMismatch = 13
    FileNotFound = 53
    FileAlreadyOpen = 55
    FileAlreadyExists = 58
    DiskFull = 61
    PermissionDenied = 70
    PathNotFound = 76
    InvalidObjectReference = 91
    
    InvalidPropertyValue = 380
    InvalidPropertyArrayIndex = 381
    
    ' non-generic error codes
    InvalidTickerID = vbObjectError + 512
    NotReceivingMarketDepth
    TickfileReplayProhibitsLiveOrders
    UnexpectedContract
    UnknownOrderID
    ContractDetailsReqNotAllowed
    InvalidOrderType
    InvalidOrderTypeInThisContext
    ContractCannotBeParsed
    NoInputTickFile
    CantCreateCrescendoTickfile
    ErrorOpeningTickfile
    CantAccessCrescendoDB
    NotImplemented
    AttemptToUseDeadTickerObject
    CantAddColumn
    CantGenerateColumns
    ColumnAlreadyAdded
    ColumnNameNotUnique
    TickerAlreadyInUse
    AlreadyConnected
    NoContractOrTickfile
    DatasourceHasNotBeenGenerated
    CantCreateGUID
    NotCorrectServiceProviderType
    NotUniqueServiceProviderName
    ServiceProviderNameInvalid
    InvalidServiceProviderHandle
    UnknownOrderTypeFromTWS

    ' generic error codes
    ArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ClassCastException              ' attempt to cast an object to class of which it is not an instance
    IllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    IllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    IndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    NullPointerException            ' attempt to use Nothing in a case where an object is required
    NumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    RuntimeException                ' an unspecified runtime error has occurred
    SecurityException               ' a security violation has occurred
    UnsupportedOperationException   ' the requested operation is not supported



End Enum

Public Enum TickTypes
    Bid
    Ask
    closePrice
    highPrice
    lowPrice
    marketDepth
    MarketDepthReset
    Trade
    volume
    openInterest
End Enum

'================================================================================
' Types
'================================================================================

Private Type TWSAPITableEntry
    server          As String
    port            As Long
    clientID        As Long
    connectionRetryIntervalSecs As Long
    keepConnection  As Boolean  ' once this flag is set, the TWSAPI instance
                                ' will only be disconnected by a call to
                                ' gReleaseTWSAPIInstance with <forceDisconnect>
                                ' set to true or by a call to
                                ' gReleaseAllTWSAPIInstances
    TWSAPI          As TWSAPI
    usageCount      As Long
End Type

'================================================================================
' Global variables
'================================================================================

Public gTWSAPI As TWSAPI

'================================================================================
' Private variables
'================================================================================

Private mTWSAPITable() As TWSAPITableEntry
Private mTWSAPITableNextIndex As Long

'================================================================================
' Procedures
'================================================================================

Public Function gCurrentTime() As Date
gCurrentTime = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
End Function

Public Function gGetTWSAPIInstance( _
                ByVal server As String, _
                ByVal port As Long, _
                ByVal clientID As Long, _
                ByVal connectionRetryIntervalSecs As Long, _
                ByVal keepConnection As Boolean) As TWSAPI
Dim i As Long

If mTWSAPITableNextIndex = 0 Then
    ReDim mTWSAPITable(5) As TWSAPITableEntry
End If

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).server = server And _
        mTWSAPITable(i).port = port And _
        mTWSAPITable(i).clientID = clientID And _
        mTWSAPITable(i).connectionRetryIntervalSecs = connectionRetryIntervalSecs _
    Then
        Set gGetTWSAPIInstance = mTWSAPITable(i).TWSAPI
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount + 1
        If keepConnection Then mTWSAPITable(i).keepConnection = True
        Exit Function
    End If
Next

If mTWSAPITableNextIndex > UBound(mTWSAPITable) Then
    ReDim Preserve mTWSAPITable(UBound(mTWSAPITable) + 5) As TWSAPITableEntry
End If

mTWSAPITable(mTWSAPITableNextIndex).server = server
mTWSAPITable(mTWSAPITableNextIndex).port = port
mTWSAPITable(mTWSAPITableNextIndex).clientID = clientID
mTWSAPITable(mTWSAPITableNextIndex).connectionRetryIntervalSecs = connectionRetryIntervalSecs
mTWSAPITable(mTWSAPITableNextIndex).usageCount = 1
Set mTWSAPITable(mTWSAPITableNextIndex).TWSAPI = New TWSAPI
Set gGetTWSAPIInstance = mTWSAPITable(mTWSAPITableNextIndex).TWSAPI

mTWSAPITableNextIndex = mTWSAPITableNextIndex + 1

gGetTWSAPIInstance.server = server
gGetTWSAPIInstance.port = port
gGetTWSAPIInstance.clientID = clientID
gGetTWSAPIInstance.connectionRetryIntervalSecs = connectionRetryIntervalSecs
gGetTWSAPIInstance.Connect

End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = 0
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
End Function

Public Sub gReleaseAllTWSAPIInstances()

Dim i As Long

For i = 0 To mTWSAPITableNextIndex - 1
    mTWSAPITable(i).usageCount = 0
    If Not mTWSAPITable(i).TWSAPI Is Nothing Then
        mTWSAPITable(i).TWSAPI.disconnect
        Set mTWSAPITable(i).TWSAPI = Nothing
    End If
    mTWSAPITable(i).clientID = 0
    mTWSAPITable(i).connectionRetryIntervalSecs = 0
    mTWSAPITable(i).port = 0
    mTWSAPITable(i).server = ""
Next
                
End Sub

Public Sub gReleaseTWSAPIInstance( _
                ByVal instance As TWSAPI, _
                Optional ByVal forceDisconnect As Boolean)

Dim i As Long

For i = 0 To mTWSAPITableNextIndex - 1
    If mTWSAPITable(i).TWSAPI Is instance Then
        mTWSAPITable(i).usageCount = mTWSAPITable(i).usageCount - 1
        If mTWSAPITable(i).usageCount = 0 And _
            ((Not mTWSAPITable(i).keepConnection) Or _
                forceDisconnect) _
        Then
            mTWSAPITable(i).TWSAPI.disconnect
            Set mTWSAPITable(i).TWSAPI = Nothing
            mTWSAPITable(i).clientID = 0
            mTWSAPITable(i).connectionRetryIntervalSecs = 0
            mTWSAPITable(i).port = 0
            mTWSAPITable(i).server = ""
        End If
        Exit For
    End If
Next
                
End Sub
                
Public Function LegOpenCloseFromString(ByVal value As String) As LegOpenClose
Select Case UCase$(value)
Case ""
    LegOpenCloseFromString = LegUnknownPos
Case "SAME"
    LegOpenCloseFromString = LegSamePos
Case "OPEN"
    LegOpenCloseFromString = LegOpenPos
Case "CLOSE"
    LegOpenCloseFromString = LegClosePos
End Select
End Function

Public Function LegOpenCloseToString(ByVal value As LegOpenClose) As String
Select Case value
Case LegSamePos
    LegOpenCloseToString = "SAME"
Case LegOpenPos
    LegOpenCloseToString = "OPEN"
Case LegClosePos
    LegOpenCloseToString = "CLOSE"
End Select
End Function

Public Function optRightFromString(ByVal value As String) As OptionRights
Select Case UCase$(value)
Case ""
    optRightFromString = OptNone
Case "CALL"
    optRightFromString = OptCall
Case "PUT"
    optRightFromString = OptPut
End Select
End Function

Public Function optRightToString(ByVal value As OptionRights) As String
Select Case value
Case OptNone
    optRightToString = ""
Case OptCall
    optRightToString = "CALL"
Case OptPut
    optRightToString = "PUT"
End Select
End Function

Public Function orderActionFromString(ByVal value As String) As OrderActions
Select Case UCase$(value)
Case "BUY"
    orderActionFromString = OrderActions.ActionBuy
Case "SELL"
    orderActionFromString = OrderActions.ActionSell
End Select
End Function

Public Function orderActionToString(ByVal value As OrderActions) As String
Select Case value
Case OrderActions.ActionBuy
    orderActionToString = "BUY"
Case OrderActions.ActionSell
    orderActionToString = "SELL"
End Select
End Function

Public Function secTypeFromString(ByVal value As String) As SecurityTypes
Select Case UCase$(value)
Case "STK"
    secTypeFromString = SecTypeStock
Case "FUT"
    secTypeFromString = SecTypeFuture
Case "OPT"
    secTypeFromString = SecTypeOption
Case "FOP"
    secTypeFromString = SecTypeFuturesOption
Case "CASH"
    secTypeFromString = SecTypeCash
Case "IND"
    secTypeFromString = SecTypeIndex
End Select
End Function

Public Function secTypeToString(ByVal value As SecurityTypes) As String
Select Case value
Case SecTypeStock
    secTypeToString = "STK"
Case SecTypeFuture
    secTypeToString = "FUT"
Case SecTypeOption
    secTypeToString = "OPT"
Case SecTypeFuturesOption
    secTypeToString = "FOP"
Case SecTypeCash
    secTypeToString = "CASH"
Case SecTypeIndex
    secTypeToString = "IND"
End Select
End Function

Public Function gTruncateTimeToNextMinute(ByVal timestamp As Date) As Date
gTruncateTimeToNextMinute = Int((timestamp + OneMinute - OneMicrosecond) / OneMinute) * OneMinute
End Function

Public Function gTruncateTimeToMinute(ByVal timestamp As Date) As Date
gTruncateTimeToMinute = Int((timestamp + OneMicrosecond) / OneMinute) * OneMinute
End Function


