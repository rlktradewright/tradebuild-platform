VERSION 5.00
Begin VB.UserControl ContractsConfigurer 
   ClientHeight    =   4305
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7800
   ScaleHeight     =   4305
   ScaleWidth      =   7800
   Begin VB.CommandButton RemoveButton 
      Caption         =   "X"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      ToolTipText     =   "Delete"
      Top             =   3240
      Width           =   375
   End
   Begin VB.CommandButton UpButton 
      Caption         =   "ñ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Picture         =   "ContractsConfigurer.ctx":0000
      TabIndex        =   3
      ToolTipText     =   "Move up"
      Top             =   1440
      Width           =   375
   End
   Begin VB.CommandButton DownButton 
      Caption         =   "ò"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   18
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Picture         =   "ContractsConfigurer.ctx":0442
      TabIndex        =   2
      ToolTipText     =   "Move down"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6960
      TabIndex        =   1
      ToolTipText     =   "Add new"
      Top             =   120
      Width           =   375
   End
   Begin VB.ListBox ContractsList 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.Shape OutlineBox 
      Height          =   4000
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "ContractsConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

