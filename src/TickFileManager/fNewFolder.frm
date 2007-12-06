VERSION 5.00
Begin VB.Form fNewFolder 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dialog Caption"
   ClientHeight    =   1245
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox NewFolderText 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
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
   Begin VB.Label Label1 
      Caption         =   "New folder name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "fNewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
NewFolderText = ""
Me.Hide
End Sub

Private Sub OKButton_Click()
Me.Hide
End Sub
