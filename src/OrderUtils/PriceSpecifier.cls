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

Private mInitialised                                As Boolean
Private mPrice                                      As Double
Private mPriceString                                As String
Private mPriceType                                  As PriceValueTypes
Private mOffset                                     As Double
Private mOffsetType                                 As PriceOffsetTypes
Private mUseCloseoutSemantics                       As Boolean
Private mTickSize                                   As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mPrice = MaxDouble
End Sub

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
Assert Not mInitialised, "Already set via initialiser"
mPrice = Value
End Property

Public Property Get PriceString() As String
PriceString = mPriceString
End Property

Public Property Let PriceString(ByVal Value As String)
Assert Not mInitialised, "Already set via initialiser"
mPriceString = Value
End Property

Public Property Get Price() As Double
Price = mPrice
End Property

Public Property Let PriceType(ByVal Value As PriceValueTypes)
Assert Not mInitialised, "Already set via initialiser"
mPriceType = Value
End Property

Public Property Get PriceType() As PriceValueTypes
PriceType = mPriceType
End Property

Public Property Let Offset(ByVal Value As Double)
Assert Not mInitialised, "Already set via initialiser"
mOffset = Value
End Property

Public Property Get Offset() As Double
Offset = mOffset
End Property

Public Property Let OffsetType(ByVal Value As PriceOffsetTypes)
Assert Not mInitialised, "Already set via initialiser"
mOffsetType = Value
End Property

Public Property Get OffsetType() As PriceOffsetTypes
OffsetType = mOffsetType
End Property

Public Property Let UseCloseoutSemantics(ByVal Value As Boolean)
Assert Not mInitialised, "Already set via initialiser"
mUseCloseoutSemantics = Value
End Property

Public Property Get UseCloseoutSemantics() As Boolean
UseCloseoutSemantics = mUseCloseoutSemantics
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub CheckPriceValid( _
                ByVal pTickSize As Double, _
                ByVal pSecType As SecurityTypes)
Const ProcName As String = "CheckPriceValid"
On Error GoTo Err

If mPriceType = PriceValueTypeValue And _
    mTickSize = -1 _
Then
    ' we only get here if this PriceSpecifier was set before the
    ' minimum tick had been determined
    mTickSize = pTickSize
    Dim lPrice As Double
    AssertArgument ParsePrice(mPriceString, pSecType, mTickSize, lPrice), "Price not valid for this tick size"
End If

Exit Sub

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Clone() As PriceSpecifier
Set Clone = New PriceSpecifier
Clone.Initialise mPrice, mPriceString, mPriceType, mOffset, mOffsetType, mTickSize, mUseCloseoutSemantics
End Function

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
                ByVal pPriceString As String, _
                ByVal pPriceType As PriceValueTypes, _
                ByVal pOffset As Double, _
                ByVal pOffsetType As PriceOffsetTypes, _
                ByVal pTickSize As Double, _
                ByVal pUseCloseoutSemantics As Boolean)
mInitialised = True
mPrice = pPrice
mPriceString = pPriceString
mPriceType = pPriceType
mOffset = pOffset
mOffsetType = pOffsetType
mTickSize = pTickSize
mUseCloseoutSemantics = pUseCloseoutSemantics
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
jb.EncodeNameAndValue "PriceString", mPriceString
jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
GOrders.HandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




