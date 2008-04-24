VERSION 5.00
Object = "{793BAAB8-EDA6-4810-B906-E319136FDF31}#62.0#0"; "TradeBuildUI2-6.ocx"
Begin VB.Form fContractSpec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contract specifier"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin TradeBuildUI26.ContractSpecBuilder ContractSpecBuilder1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   5106
   End
End
Attribute VB_Name = "fContractSpec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

