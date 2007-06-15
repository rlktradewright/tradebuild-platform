VERSION 5.00
Begin VB.Form fTimeframeSpecifier 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Specify a timeframe"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TradeBuildUI26.TimeframeSpecifier TimeframeSpecifier1 
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   600
      Width           =   735
   End
   Begin VB.CommandButton OkButton 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "fTimeframeSpecifier"
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

Private Const ProjectName As String = "TradeBuildUI25"
Private Const ModuleName As String = "fTimeframeSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Activate()
mCancelled = False
TimeframeSpecifier1.SetFocus
End Sub

Private Sub Form_Load()
If TimeframeSpecifier1.isTimeframeValid Then OkButton.Enabled = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
mCancelled = True
Me.Hide
End Sub

Private Sub OkButton_Click()
Me.Hide
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Friend Property Get timeframeDesignator() As TimePeriod
Dim tp As TimePeriod
If mCancelled Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & "timeframeDesignator", _
                "Cancelled by user"
End If
tp = TimeframeSpecifier1.timeframeDesignator
timeframeDesignator = tp
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


Private Sub TimeframeSpecifier1_Change()
If TimeframeSpecifier1.isTimeframeValid Then
    OkButton.Enabled = True
Else
    OkButton.Enabled = False
End If
End Sub
