Attribute VB_Name = "Globals"
Option Explicit

'================================================================================
' Constants
'================================================================================

Public Const OneMicrosecond As Double = 1# / 86400000000#
Public Const OneMinute As Double = 1# / 1440#

Public Const ContractInfoSPName As String = "IB TWS Contract Info Service Provider"
Public Const HistoricDataSPName As String = "IB TWS Historic Data Service Provider"
Public Const RealtimeDataSPName As String = "IB TWS Realtime Data Service Provider"
Public Const OrderSubmissionSPName As String = "IB TWS Order Submission Service Provider"

Public Const providerKey As String = "TWS"

'================================================================================
' Enums
'================================================================================

Public Enum ErrorCodes
    ' generic run-time error codes defined by VB
    ErrInvalidProcedureCall = 5
    ErrOverflow = 6
    ErrSubscriptOutOfRange = 9
    ErrDivisionByZero = 11
    ErrTypeMismatch = 13
    ErrFileNotFound = 53
    ErrFileAlreadyOpen = 55
    ErrFileAlreadyExists = 58
    ErrDiskFull = 61
    ErrPermissionDenied = 70
    ErrPathNotFound = 76
    ErrInvalidObjectReference = 91
    
    ErrInvalidPropertyValue = 380
    ErrInvalidPropertyArrayIndex = 381
    
    ' generic error codes
    ErrArithmeticException = vbObjectError + 1024  ' an exceptional arithmetic condition has occurred
    ErrArrayIndexOutOfBoundsException  ' an array has been accessed with an illegal index
    ErrClassCastException              ' attempt to cast an object to class of which it is not an instance
    ErrIllegalArgumentException        ' method has been passed an illegal or inappropriate argument
    ErrIllegalStateException           ' a method has been invoked at an illegal or inappropriate time
    ErrIndexOutOfBoundsException       ' an index of some sort (such as to an array, to a string, or to a vector) is out of range
    ErrNullPointerException            ' attempt to use Nothing in a case where an object is required
    ErrNumberFormatException           ' attempt to convert a string to one of the numeric types, but the string does not have the appropriate format
    ErrRuntimeException                ' an unspecified runtime error has occurred
    ErrSecurityException               ' a security violation has occurred
    ErrUnsupportedOperationException   ' the requested operation is not supported
End Enum

'================================================================================
' Types
'================================================================================

Private Type TWSAPITableEntry
    server          As String
    port            As Long
    clientID        As Long
    providerKey     As String
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

'================================================================================
' Private variables
'================================================================================

Private mCommonServiceConsumer As ICommonServiceConsumer
Private mTWSAPITable() As TWSAPITableEntry
Private mTWSAPITableNextIndex As Long

'================================================================================
' Procedures
'================================================================================

Public Property Let gCommonServiceConsumer( _
                ByVal RHS As TradeBuildSP.ICommonServiceConsumer)
Set mCommonServiceConsumer = RHS
End Property

Public Function gCurrentTime() As Date
gCurrentTime = CDbl(Int(Now)) + (CDbl(Timer) / 86400#)
End Function

Public Function gGetTWSAPIInstance( _
                ByVal server As String, _
                ByVal port As Long, _
                ByVal clientID As Long, _
                ByVal providerKey As String, _
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
        mTWSAPITable(i).providerKey = providerKey And _
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
mTWSAPITable(mTWSAPITableNextIndex).providerKey = providerKey
mTWSAPITable(mTWSAPITableNextIndex).connectionRetryIntervalSecs = connectionRetryIntervalSecs
mTWSAPITable(mTWSAPITableNextIndex).usageCount = 1
Set mTWSAPITable(mTWSAPITableNextIndex).TWSAPI = New TWSAPI
Set gGetTWSAPIInstance = mTWSAPITable(mTWSAPITableNextIndex).TWSAPI

mTWSAPITableNextIndex = mTWSAPITableNextIndex + 1

gGetTWSAPIInstance.commonServiceConsumer = mCommonServiceConsumer
gGetTWSAPIInstance.server = server
gGetTWSAPIInstance.port = port
gGetTWSAPIInstance.clientID = clientID
gGetTWSAPIInstance.providerKey = providerKey
gGetTWSAPIInstance.connectionRetryIntervalSecs = connectionRetryIntervalSecs
gGetTWSAPIInstance.Connect

End Function

Public Function gHistDataCapabilities() As Long
gHistDataCapabilities = 0
End Function

Public Function gHistDataSupports(ByVal capabilities As Long) As Boolean
gHistDataSupports = (gHistDataCapabilities And capabilities)
End Function

Public Function gRealtimeDataCapabilities() As Long
gRealtimeDataCapabilities = TradeBuildSP.RealtimeDataServiceProviderCapabilities.RtCapMarketDepthByPosition
End Function

Public Function gRealtimeDataSupports(ByVal capabilities As Long) As Boolean
gRealtimeDataSupports = (gRealtimeDataCapabilities And capabilities)
End Function

Public Sub gReleaseAllTWSAPIInstances()

Dim i As Long

For i = 0 To mTWSAPITableNextIndex - 1
    mTWSAPITable(i).usageCount = 0
    If Not mTWSAPITable(i).TWSAPI Is Nothing Then
        mTWSAPITable(i).TWSAPI.disconnect "release all"
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
            mTWSAPITable(i).TWSAPI.disconnect "release"
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


