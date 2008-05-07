VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bars"
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

Implements CollectionChangeListener
Implements Enumerable

'@================================================================================
' Events
'@================================================================================

''
' Raised when a bar is added to or removed from the
' collection, and when the collection is cleared.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>CollectionChangedEvent</code> parameter:
' <blockquote>
'    <tt>CollItemAdded</tt><br />
' </blockquote>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "Bars"

'@================================================================================
' Member variables
'@================================================================================

Private mColl                               As EnumerableCollection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mColl = New EnumerableCollection
mColl.addCollectionChangeListener Me
End Sub

'@================================================================================
' CollectionChangeListener Interface Members
'@================================================================================

Private Sub CollectionChangeListener_Change( _
                ev As TWUtilities30.CollectionChangeEvent)
RaiseEvent CollectionChanged(ev)
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As TWUtilities30.Enumerator
Set Enumerable_Enumerator = mColl.Enumerator
End Property

Private Function Enumerable_nextItem( _
                item As Variant, _
                cookie As Variant) As Boolean

End Function

Private Function Enumerable_skip( _
                ByVal numToSkip As Long, _
                cookie As Variant) As Boolean

End Function

Private Property Get Enumerable_updateSequence() As Long

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Returns an code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Set Enumerator = mColl.Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for changes to the collection.
'
' @remarks
'   The listener object must implement the <code>CollectionChangeListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub addCollectionChangeListener( _
                ByVal listener As CollectionChangeListener)
mColl.addCollectionChangeListener listener
End Sub

''
' Returns the number of items in the collection.
'
' @return
'   The number of items in the collection.
'@/
Public Function count() As Long
count = mColl.count
End Function

''
' Returns a specific Bar in the collection.
'
' @remarks
'   If the value provided as index doesn’t match any existing member of the
'   collection, an error occurs.
'
' @return
'   The required Bar.
' @param index
'   An expression that specifies the position of a member of the collection.
'   If a numeric expression, index must be a number from 1 to the value of the
'   collection's Count property. If a string expression, index must
'   correspond to the key argument specified when the member referred to was
'   added to the collection.
'@/
Public Function item( _
                ByRef index As Variant) As Bar
Dim v As Variant
Bar = mColl.item(index)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


