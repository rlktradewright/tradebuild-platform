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

Private Const ModuleName As String = "fTimeframeSpecifier"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Activate()
Const ProcName As String = "Form_Activate"
Dim failpoint As String
On Error GoTo Err

mCancelled = False
TimeframeSpecifier1.SetFocus

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Load()
Const ProcName As String = "Form_Load"
Dim failpoint As String
On Error GoTo Err

If TimeframeSpecifier1.isTimeframeValid Then OkButton.Enabled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
Dim failpoint As String
On Error GoTo Err

mCancelled = True
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"
Dim failpoint As String
On Error GoTo Err

Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' TimeframeSpecifier1 Event Handlers
'@================================================================================

Private Sub TimeframeSpecifier1_Change()
Const ProcName As String = "TimeframeSpecifier1_Change"
Dim failpoint As String
On Error GoTo Err

If TimeframeSpecifier1.isTimeframeValid Then
    OkButton.Enabled = True
Else
    OkButton.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Friend Property Get TimeframeDesignator() As TimePeriod
Const ProcName As String = "timeframeDesignator"
Dim failpoint As String
On Error GoTo Err

If mCancelled Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Cancelled by user"
End If
Set TimeframeDesignator = TimeframeSpecifier1.TimeframeDesignator

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


