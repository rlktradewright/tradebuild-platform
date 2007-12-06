VERSION 5.00
Begin VB.Form fPathChooser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Choose folder"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton NewFolderButton 
      Cancel          =   -1  'True
      Caption         =   "New folder..."
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.DirListBox DirList 
      Height          =   2565
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4335
   End
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4335
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

Private mCancelled As Boolean

Public Property Let path(ByVal newvalue As String)
If Mid$(newvalue, 2, 1) = ":" Then
    DriveList.Drive = Left$(newvalue, 2)
    DirList.path = newvalue
ElseIf Left$(newvalue, 2) = "\\" Then
End If

End Property

Public Property Get path() As String
If Not mCancelled Then path = DirList.path
End Property

Private Sub CancelButton_Click()
Me.Hide
mCancelled = True
End Sub

Private Sub DriveList_Change()
DirList.path = DriveList.Drive
End Sub

Private Sub NewFolderButton_Click()
Dim fNew As New fNewFolder
Dim filesys As FileSystemObject
Dim folder As folder
Dim folders As folders
Dim newFolderPath As String

On Error GoTo err

show:

fNew.show vbModal
If fNew.NewFolderText = "" Then Unload fNew: Exit Sub

Set filesys = New FileSystemObject
Set folder = filesys.GetFolder(DirList.path)
Set folders = folder.SubFolders
folders.Add fNew.NewFolderText

newFolderPath = DirList.path & "\" & fNew.NewFolderText

DirList.Refresh
DirList.path = newFolderPath
Unload fNew
Exit Sub

err:
If err.Number = 58 Then
    ' folder already exists
    MsgBox "Folder already exists", , "Error"
    Resume show
End If
Unload fNew
End Sub

Private Sub OKButton_Click()
mCancelled = False
Me.Hide
End Sub
