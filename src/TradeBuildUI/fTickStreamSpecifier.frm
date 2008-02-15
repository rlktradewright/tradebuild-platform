VERSION 5.00
Begin VB.Form fTickStreamSpecifier 
   Caption         =   "Tickstream specifier"
   ClientHeight    =   3930
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6840
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   6840
      TabIndex        =   2
      Top             =   840
      Width           =   735
   End
   Begin TradeBuildUI26.TickStreamSpecifier TickStreamSpecifier1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6588
   End
End
Attribute VB_Name = "fTickStreamSpecifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
' @remarks
' @see
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
Private Const ModuleName                    As String = "fTickstreamSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mTickfileSpecifiers()               As TickfileSpecifier

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Load()
mCancelled = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
mCancelled = True
Unload Me
End Sub

Private Sub OKButton_Click()
Screen.MousePointer = vbHourglass
mCancelled = False
TickStreamSpecifier1.load
End Sub

Private Sub TickStreamSpecifier1_NotReady()
OkButton.Enabled = False
End Sub

Private Sub TickStreamSpecifier1_ready()
OkButton.Enabled = True
End Sub

Private Sub TickStreamSpecifier1_TickStreamsSpecified( _
                pTickfileSpecifiers() As TickfileSpecifier)
Screen.MousePointer = vbDefault
mTickfileSpecifiers = pTickfileSpecifiers
Me.Hide
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Get TickfileSpecifiers() As TickfileSpecifier()
TickfileSpecifiers = mTickfileSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================





