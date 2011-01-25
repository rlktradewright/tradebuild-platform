VERSION 5.00
Object = "{7837218F-7821-47AD-98B6-A35D4D3C0C38}#48.0#0"; "TWControls10.ocx"
Begin VB.UserControl ChartStylePicker 
   ClientHeight    =   315
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1935
   ScaleHeight     =   315
   ScaleWidth      =   1935
   Begin TWControls10.TWImageCombo ChartStylesCombo 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "ChartStylePicker.ctx":0000
      Text            =   ""
   End
End
Attribute VB_Name = "ChartStylePicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

