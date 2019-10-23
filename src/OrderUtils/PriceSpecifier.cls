VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PriceSpecifier"
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

Implements IJSONable

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

Private Const ModuleName                            As String = "PriceSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mPrice                                      As Double
Private mPriceType                                  As PriceValueTypes
Private mOffset                                     As Double
Private mOffsetType                                 As PriceOffsetTypes

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJSONAble_ToJSON() As String
IJSONAble_ToJSON = ToJSON
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get IsValid() As Boolean
IsValid = mPriceType <> PriceValueTypeNone
End Property

Public Property Let Price(ByVal Value As Double)
mPrice = Value
End Property

Public Property Get Price() As Double
Price = mPrice
End Property

Public Property Let PriceType(ByVal Value As PriceValueTypes)
mPriceType = Value
End Property

Public Property Get PriceType() As PriceValueTypes
PriceType = mPriceType
End Property

Public Property Let Offset(ByVal Value As Double)
mOffset = Value
End Property

Public Property Get Offset() As Double
Offset = mOffset
End Property

Public Property Let OffsetType(ByVal Value As PriceOffsetTypes)
mOffsetType = Value
End Property

Public Property Get OffsetType() As PriceOffsetTypes
OffsetType = mOffsetType
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Equals(ByVal pObj As Object) As Boolean
If pObj Is Nothing Then Exit Function
If Not TypeOf pObj Is PriceSpecifier Then Exit Function
If Not pObj Is Me Then
    Dim lPriceSpec As PriceSpecifier: Set lPriceSpec = pObj
    If lPriceSpec.Price <> mPrice Then Exit Function
    If lPriceSpec.PriceType <> mPriceType Then Exit Function
    If lPriceSpec.Offset <> mOffset Then Exit Function
    If lPriceSpec.OffsetType <> mOffsetType Then Exit Function
End If
Equals = True
End Function

Friend Sub Initialise( _
                ByVal pPrice As Double, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes)
mPrice = pPrice
mPriceType = pPriceType
mOffset = pOffset
mOffsetType = pOffsetType
End Sub

Public Function RequiresDelayedPriceResolution() As Boolean
RequiresDelayedPriceResolution = (mPriceType = PriceValueTypeEntry)
End Function

Friend Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeNameAndValue "Price", mPrice, MaxDouble
jb.EncodeNameAndValue "Offset", mOffset, 0#
jb.EncodeNameAndValue "PriceType", mPriceType, PriceValueTypes.PriceValueTypeNone
jb.EncodeNameAndValue "OffsetType", mOffsetType, PriceOffsetTypes.PriceOffsetTypeNone

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


'@================================================================================
' Helper Functions
'@================================================================================



