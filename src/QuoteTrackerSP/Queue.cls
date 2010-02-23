VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const S_OK = 0

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type GUID
    data(4) As Long
End Type

Private Type GUIDString
    StartBrace  As String * 2
    GUIDProper  As String * 72
    EndBrace    As String * 2
    ZeroByte    As String * 1
End Type

Private Type QueueEntry
    item                As Variant
    itemKey             As String
    itemRemoved         As Boolean
End Type

'================================================================================
' External Function Declarations
'================================================================================

Private Declare Function CoCreateGuid Lib "OLE32.dll" (pGUID As GUID) As Long

Private Declare Function StringFromGUID2 Lib "OLE32.dll" ( _
                            ByRef rguid As GUID, _
                            ByRef lpsz As GUIDString, _
                            ByVal cchMax As Long) As Integer

'================================================================================
' Member variables
'================================================================================

Private mInitialised As Boolean

Private mEntries() As QueueEntry
Private mSize As Long
Private mFirstEntry As Long
Private mNextEntry As Long

Private mCapacityIncrement As Long

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Friend Property Get capacity() As Long
capacity = UBound(mEntries) + 1
End Property

Friend Property Let capacityIncrement(ByVal value As Long)
If value < 1 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "IBTWSSP.Queue::capacityIncrement", _
                            "Capaciy increment must be >0"
mCapacityIncrement = value
End Property

Friend Property Get frontKey() As String
If mSize = 0 Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                "Queue::frontKey", _
                                "Queue empty"

Do While mEntries(mFirstEntry).itemRemoved
    mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
Loop

frontKey = mEntries(mFirstEntry).itemKey
End Property

Friend Property Let initialCapacity(ByVal value As Long)
If value < 1 Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                            "IBTWSSP.Queue::initialCapacity", _
                            "Initial capaciy must be >0"
ReDim mEntries(value - 1) As QueueEntry
If mCapacityIncrement = 0 Then mCapacityIncrement = value
mInitialised = True
End Property

Friend Property Get size() As Long
size = mSize
End Property

'================================================================================
' Methods
'================================================================================

Friend Function addToEnd(ByVal Entry As Variant) As String
If Not mInitialised Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                "Queue::addToEnd", _
                                "Queue not initialised"
If mNextEntry = mFirstEntry And mSize <> 0 Then resize
addToEnd = generateGUIDString
mEntries(mNextEntry).item = Entry
mEntries(mNextEntry).itemKey = addToEnd
mNextEntry = (mNextEntry + 1) Mod (UBound(mEntries) + 1)
mSize = mSize + 1
End Function

Friend Sub remove(ByVal itemRef As Variant)
Dim i As Long
Dim maxIndex As Long

If mSize = 0 Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                "Queue::remove", _
                                "Queue empty"

If mFirstEntry < mNextEntry Then
    maxIndex = mNextEntry - 1
Else
    maxIndex = UBound(mEntries) + mNextEntry
End If

If VarType(itemRef) = vbString Then
    For i = mFirstEntry To maxIndex
        If mEntries(i Mod (UBound(mEntries) - 1)).itemKey = itemRef Then
            mEntries(i Mod (UBound(mEntries) - 1)).itemRemoved = True
            mEntries(i Mod (UBound(mEntries) - 1)).itemKey = ""
            mEntries(i Mod (UBound(mEntries) - 1)).item = Empty
            Exit For
        End If
    Next
ElseIf VarType(itemRef) = vbObject Then
    For i = mFirstEntry To maxIndex
        If mEntries(i Mod (UBound(mEntries) - 1)).item Is itemRef Then
            mEntries(i Mod (UBound(mEntries) - 1)).itemRemoved = True
            mEntries(i Mod (UBound(mEntries) - 1)).itemKey = ""
            mEntries(i Mod (UBound(mEntries) - 1)).item = Empty
            Exit For
        End If
    Next
Else
    err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "Queue::remove", _
                "Argument type must be string or object"
End If

If i > maxIndex Then err.Raise ErrorCodes.ErrIllegalArgumentException, _
                                "Queue::remove", _
                                "Specified item not in queue"

mSize = mSize - 1
End Sub

Friend Function removeFromFront() As Variant
Dim failpoint As Long

On Error GoTo err

If mSize = 0 Then err.Raise ErrorCodes.ErrIllegalStateException, _
                                "Queue::removeFromFront", _
                                "Queue empty"

Do While mEntries(mFirstEntry).itemRemoved
    mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
Loop

failpoint = 100
removeFromFront = mEntries(mFirstEntry).item

failpoint = 200
mEntries(mFirstEntry).item = Empty
mEntries(mFirstEntry).itemKey = ""
mFirstEntry = (mFirstEntry + 1) Mod (UBound(mEntries) + 1)
mSize = mSize - 1

Exit Function

err:
err.Raise err.Number, , "Queue::removeFromFront(failpoint=" & failpoint & "): " & err.Description
End Function

'================================================================================
' Helper Functions
'================================================================================

Private Function generateGUID() As GUID
Dim lReturn As Long

lReturn = CoCreateGuid(generateGUID)

If (lReturn <> S_OK) Then
    err.Raise ErrorCodes.ErrRuntimeException, _
                "Queue::generateGUID", _
                "Can't create GUID"
End If

End Function

Private Function generateGUIDString() As String
generateGUIDString = GUIDToString(generateGUID)
End Function

Private Function GUIDToString(ByRef pGUID As GUID) As String
Dim GUIDString As GUIDString
Dim iChars As Integer

' convert binary GUID to string form
iChars = StringFromGUID2(pGUID, GUIDString, Len(GUIDString))
' convert string to ANSI
GUIDToString = StrConv(GUIDString.GUIDProper, vbFromUnicode)
End Function

Private Sub resize()
Dim i As Long
Dim currUbound As Long

currUbound = UBound(mEntries)

ReDim Preserve mEntries(currUbound + mCapacityIncrement) As QueueEntry

If mNextEntry <= mFirstEntry Then
    For i = 0 To mNextEntry - 1
        mEntries((currUbound + i + 1) Mod (UBound(mEntries) + 1)) = mEntries(i)
        mEntries(i).item = Empty
        mEntries(i).itemKey = ""
        mEntries(i).itemRemoved = False
    Next
    mNextEntry = (mNextEntry + currUbound + 1) Mod (UBound(mEntries) + 1)
End If
End Sub



