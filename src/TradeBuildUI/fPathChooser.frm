VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#40.1#0"; "TWControls10.ocx"
Begin VB.Form fPathChooser 
   Caption         =   "Choose folder"
   ClientHeight    =   2775
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6030
   Begin TWControls10.PathChooser PathChooser1 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
   End
   Begin VB.CommandButton NewFolderButton 
      Cancel          =   -1  'True
      Caption         =   "New folder..."
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "fPathChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Const ModuleName                    As String = "fPathChooser"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Activate()
mCancelled = True
End Sub

Private Sub Form_Initialize()
mCancelled = True
End Sub

Private Sub Form_Resize()
Dim butleft As Long
Const ProcName As String = "Form_Resize"
Dim failpoint As String
On Error GoTo Err

butleft = Me.ScaleWidth - OkButton.Width - 8 * Screen.TwipsPerPixelX
If butleft >= 2160 Then
    OkButton.Left = butleft
    CancelButton.Left = butleft
    NewFolderButton.Left = butleft
    PathChooser1.Width = butleft - 8 * Screen.TwipsPerPixelX - PathChooser1.Left
End If

If Me.ScaleHeight >= 1560 Then
    PathChooser1.Height = Me.ScaleHeight - 8 * Screen.TwipsPerPixelY - PathChooser1.Top
    NewFolderButton.Top = PathChooser1.Height + PathChooser1.Top - NewFolderButton.Height
End If

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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

Me.Hide
mCancelled = True

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub NewFolderButton_Click()
Const ProcName As String = "NewFolderButton_Click"
Dim failpoint As String
On Error GoTo Err

PathChooser1.NewFolder

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"
Dim failpoint As String
On Error GoTo Err

mCancelled = False
Me.Hide

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
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

Public Property Let path(ByVal newvalue As String)
Const ProcName As String = "path"
Dim failpoint As String
On Error GoTo Err

PathChooser1.path = newvalue

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

Public Property Get path() As String
Const ProcName As String = "path"
Dim failpoint As String
On Error GoTo Err

If Not mCancelled Then path = PathChooser1.path

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pProjectName:=ProjectName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


