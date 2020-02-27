VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "OPI"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestamp                                  As Date
Private mRecoveryContextName                        As String
Private mData                                       As IJSONable
Private mItemType                                   As String

Private mToString                                   As String

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

Public Property Let Data(ByVal Value As Variant)
gSetVariant mData, Value
End Property

Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

Public Property Let ItemType(ByVal Value As String)
mItemType = Value
End Property

Public Property Get ItemType() As String
ItemType = mItemType
End Property

Public Property Let RCN(ByVal Value As String)
mRecoveryContextName = Value
End Property

Public Property Get RCN() As String
RCN = mRecoveryContextName
End Property

Public Property Let Timestamp(ByVal Value As Date)
mTimestamp = Value
End Property

Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pRecoveryContextName As String, _
                ByVal pInfoType As String, _
                Optional ByVal pData As IJSONable = Nothing)
Const ProcName As String = "Initialise"
On Error GoTo Err

mTimestamp = GetTimestamp
mRecoveryContextName = pRecoveryContextName
mItemType = pInfoType
Set mData = pData

Dim s As String
s = FormatTimestamp(mTimestamp, TimestampDateAndTimeISO8601)
s = s & "," & mRecoveryContextName
s = s & "," & mItemType

If Not mData Is Nothing Then s = s & "," & """" & mData.ToJSON & """"

mToString = s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseFromString(ByVal pValue As String)
Const ProcName As String = "InitialiseFromString"
On Error GoTo Err

mToString = pValue

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(pValue, ",")

Dim lTimestampString As String: lTimestampString = lClp.Arg(0)
Dim index As Long: index = InStrRev(lTimestampString, ".")
If index = 0 Then
    mTimestamp = CDate(lTimestampString)
Else
    mTimestamp = CDate(Left$(lTimestampString, index - 1)) + _
                       CDbl(Mid$(lTimestampString, index + 1)) / 86400000
End If

mRecoveryContextName = lClp.Arg(1)
mItemType = lClp.Arg(2)

If mItemType = OrderInfoData Then
    ParseJSON lClp.Arg(3), mData
    If TypeOf mData Is PMS Then
        Dim lPMS As PMS: Set lPMS = mData
        lPMS.Diff.UpdateHost lPMS
    ElseIf TypeOf mData Is BOS Then
        Dim lBOS As BOS: Set lBOS = mData
        lBOS.Diff.UpdateHost lBOS
    ElseIf TypeOf mData Is OS Then
        Dim lOS As OS: Set lOS = mData
        lOS.Diff.UpdateHost lOS
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ToString() As String
ToString = mToString
End Function

'@================================================================================
' Helper Functions
'@================================================================================


