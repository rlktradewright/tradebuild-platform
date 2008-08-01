VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mWriter As TradeBuildSP.ITickfileWriter
Private mServiceProviderWriterReady As Boolean

Private mPendingWrites As Collection

'Private WithEvents mTimer As IntervalTimer

'Private mSelfRef As TickfileWriter

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPendingWrites = New Collection
End Sub

Private Sub Class_Terminate()
Debug.Print "TradeBuild26.TickfileWriter - object terminated"
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get numberOfPendingWrites() As Long
numberOfPendingWrites = mPendingWrites.Count
End Property

Friend Property Let ServiceProviderWriter(ByVal value As TradeBuildSP.ITickfileWriter)
Set mWriter = value
End Property

Friend Property Get ready() As Boolean
ready = mServiceProviderWriterReady
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub closeOutputFile()
Debug.Print "TradeBuild26.TickfileWriter - closing output tickfile"
mWriter.closeOutputFile
End Sub

Friend Sub ServiceProviderWriterNotReady()
mServiceProviderWriterReady = False
End Sub

Friend Sub ServiceProviderWriterReady()
Dim item As Variant
Dim pendingTick As GenericTick

mServiceProviderWriterReady = True

For Each item In mPendingWrites
    pendingTick = item
    mWriter.writeTick pendingTick
Next

Set mPendingWrites = New Collection
End Sub

Friend Sub writeTick( _
                pTick As GenericTick)
If Not mServiceProviderWriterReady Then
    mPendingWrites.add pTick
Else
    mWriter.writeTick pTick
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


