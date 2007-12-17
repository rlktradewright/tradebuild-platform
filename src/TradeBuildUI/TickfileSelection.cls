VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TickfileSelection"
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

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "TradeBuildUI26"
Private Const ModuleName                    As String = "TickfileSelection"

'@================================================================================
' Member variables
'@================================================================================

Private mTickfileOrganiser As fTickfileOrganiser

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

Public Property Get TickfileSpecifiers() As TickfileSpecifier()
If Not userCancelled Then TickfileSpecifiers = mTickfileOrganiser.TickfileSpecifiers
End Property

Public Property Get userCancelled() As Boolean
If mTickfileOrganiser Is Nothing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "userCancelled", _
            "Tickfile selection dialogue has not been displayed"
End If

userCancelled = mTickfileOrganiser.cancelled
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub SelectTickfiles()
Dim i As Long

Set mTickfileOrganiser = New fTickfileOrganiser

mTickfileOrganiser.Show vbModal, Me

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


