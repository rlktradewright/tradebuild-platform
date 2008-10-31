VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   KeyPreview      =   -1  'True
   ScaleHeight     =   7575
   ScaleWidth      =   10665
   Begin VB.PictureBox BlankPicture 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "ChartArea.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   4320
      Width           =   7455
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   600
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1138
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":158A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":19DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2280
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":26D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2B24
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":33C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":381A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":3C6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":40BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4510
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4962
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5206
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5658
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":634E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":67A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":6BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7044
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7496
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":78E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7D3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":818C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":85DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8A30
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":92D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox RegionDividerPicture 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   70
      Index           =   0
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   75
      ScaleWidth      =   9375
      TabIndex        =   2
      Top             =   6240
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.PictureBox YAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   8400
      ScaleHeight     =   615
      ScaleWidth      =   975
      TabIndex        =   3
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox XAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   9390
      TabIndex        =   1
      Top             =   6960
      Width           =   9390
   End
   Begin VB.PictureBox ChartRegionPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   0
      MouseIcon       =   "ChartArea.ctx":9726
      MousePointer    =   99  'Custom
      ScaleHeight     =   615
      ScaleWidth      =   8415
      TabIndex        =   0
      Top             =   6360
      Visible         =   0   'False
      Width           =   8415
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9B68
            Key             =   "showbars"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9E82
            Key             =   "showcandlesticks"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A19C
            Key             =   "showline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A4B6
            Key             =   "showcrosshair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A7D0
            Key             =   "showdisccursor"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":AAEA
            Key             =   "thinnerbars"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":AE04
            Key             =   "thickerbars"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B11E
            Key             =   "narrower"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B570
            Key             =   "wider"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B88A
            Key             =   "scaledown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":BBA4
            Key             =   "scaleup"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":BEBE
            Key             =   "scrolldown"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C1D8
            Key             =   "scrollup"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C4F2
            Key             =   "scrollleft"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C80C
            Key             =   "scrollright"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CB26
            Key             =   "scrollend"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CE40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D15A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D474
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D78E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DAA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E0DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E3F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E710
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":EB62
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":EFB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F2CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F5E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F902
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":FC1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":FF36
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":10250
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1056A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList4"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   22
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showbars"
            Object.ToolTipText     =   "Bar chart"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showcandlesticks"
            Object.ToolTipText     =   "Candlestick chart"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showline"
            Object.ToolTipText     =   "Line chart"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showcrosshair"
            Object.ToolTipText     =   "Show crosshair"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "showdisccursor"
            Object.ToolTipText     =   "Show cursor"
            ImageIndex      =   5
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "thinnerbars"
            Object.ToolTipText     =   "Thinner bars"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "thickerbars"
            Object.ToolTipText     =   "Thicker bars"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "reducespacing"
            Object.ToolTipText     =   "Reduce bar spacing"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "increasespacing"
            Object.ToolTipText     =   "Increase bar spacing"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scaledown"
            Object.ToolTipText     =   "Compress vertical scale"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scaleup"
            Object.ToolTipText     =   "Expand vertical scale"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scrolldown"
            Object.ToolTipText     =   "Scroll down"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scrollup"
            Object.ToolTipText     =   "Scroll up"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scrollleft"
            Object.ToolTipText     =   "Scroll left"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scrollright"
            Object.ToolTipText     =   "Scroll right"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "scrollend"
            Object.ToolTipText     =   "Scroll to end"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "autoscale"
            Object.ToolTipText     =   "Autoscale"
            ImageIndex      =   17
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

Event ChartCleared()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event RegionsChanged(ev As CollectionChangeEvent)
Event PeriodsChanged(ev As CollectionChangeEvent)
Event RegionSelected(ByVal Region As ChartRegion)

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type RegionTableEntry
    Region              As ChartRegion
    percentheight       As Double
    actualHeight        As Long
    useAvailableSpace   As Boolean
End Type

'================================================================================
' Constants
'================================================================================

Private Const ProjectName                As String = "ChartSkil26"
Private Const ModuleName                As String = "Chart"

'Private Const HorizScrollBarHeight As Long = 255
'Private Const ToolbarBarHeight As Long = 330

Private Const PropNameAllowHorizontalMouseScrolling     As String = "AllowHorizontalMouseScrolling"
Private Const PropNameAllowVerticalMouseScrolling       As String = "AllowVerticalMouseScrolling"
Private Const PropNameAutoscroll                        As String = "Autoscroll"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNameDefaultBarDisplayMode             As String = "DefaultBarDisplayMode"
Private Const PropNameDefaultRegionAutoscale            As String = "DefaultRegionAutoscale"
Private Const PropNameDefaultRegionBackColor            As String = "DefaultRegionBackColor"
Private Const PropNameDefaultRegionGridColor            As String = "DefaultRegionGridColor"
Private Const PropNameDefaultRegionGridlineSpacingY     As String = "DefaultRegionGridlineSpacingY"
Private Const PropNameDefaultRegionGridTextColor        As String = "DefaultRegionGridTextColor"
Private Const PropNameDefaultRegionHasGrid              As String = "DefaultRegionHasGrid"
Private Const PropNameDefaultRegionHasGridtext          As String = "DefaultRegionHasGridtext"
Private Const PropNameDefaultRegionIntegerYScale        As String = "DefaultRegionIntegerYScale"
Private Const PropNameDefaultRegionMinimumHeight        As String = "DefaultRegionMinimumHeight"
Private Const PropNameDefaultRegionYScaleQuantum        As String = "DefaultRegionYScaleQuantum"
Private Const PropNamePeriodLength                      As String = "PeriodLength"
Private Const PropNamePeriodUnits                       As String = "PeriodUnits"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameShowHorizontalScrollBar           As String = "ShowHorizontalScrollBar"
Private Const PropNameShowToolbar                       As String = "ShowToobar"
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
Private Const PropNameVerticalGridSpacing               As String = "VerticalGridSpacing"
Private Const PropNameVerticalGridUnits                 As String = "VerticalGridUnits"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltAllowHorizontalMouseScrolling     As Boolean = True
Private Const PropDfltAllowVerticalMouseScrolling       As Boolean = True
Private Const PropDfltAutoscroll                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = vbWhite
Private Const PropDfltDefaultBarDisplayMode             As Long = BarDisplayModes.BarDisplayModeBar
Private Const PropDfltDefaultRegionAutoscale            As Boolean = True
Private Const PropDfltDefaultRegionBackColor            As Long = vbWhite
Private Const PropDfltDefaultRegionGridColor            As Long = &HC0C0C0
Private Const PropDfltDefaultRegionGridlineSpacingY     As Double = 1.8
Private Const PropDfltDefaultRegionGridTextColor        As Long = vbBlack
Private Const PropDfltDefaultRegionHasGrid              As Boolean = True
Private Const PropDfltDefaultRegionHasGridtext          As Boolean = False
Private Const PropDfltDefaultRegionIntegerYScale        As Boolean = False
Private Const PropDfltDefaultRegionMinimumHeight        As Double = 0
Private Const PropDfltDefaultRegionYScaleQuantum        As Double = 0#
Private Const PropDfltPeriodLength                      As Long = 5
Private Const PropDfltPeriodUnits                       As Long = TimePeriodMinute
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltShowHorizontalScrollBar           As Boolean = True
Private Const PropDfltShowToolbar                       As Boolean = True
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltVerticalGridSpacing               As Long = 1
Private Const PropDfltVerticalGridUnits                 As Long = TimePeriodHour
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'================================================================================
' Member variables
'================================================================================

Private mController As ChartController

Private mRegions() As RegionTableEntry
Private mRegionsIndex As Long
Private mNumRegionsInUse As Long

Private mDefaultRegionStyle As ChartRegionStyle
Private mDefaultBarStyle As BarStyle
Private mDefaultDataPointStyle As DataPointStyle
Private mDefaultLineStyle As linestyle
Private mDefaultTextStyle As TextStyle

Private mDefaultYAxisStyle As ChartRegionStyle

Private WithEvents mPeriods As Periods
Attribute mPeriods.VB_VarHelpID = -1

'Private mAutoscale As Boolean
Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single
Private mPrevWidth As Single

Private mTwipsPerBar As Long

Private mXAxisRegion As ChartRegion
Private mXCursorText As text

Private mYAxisPosition As Long
Private mYAxisWidthCm As Single

Private mSessionStartTime As Date
Private mSessionEndTime As Date

Private mCurrentSessionStartTime As Date
Private mCurrentSessionEndTime As Date

Private mBarTimePeriod As TimePeriod
Private mBarTimePeriodSet As Boolean

Private mVerticalGridTimePeriod As TimePeriod
Private mVerticalGridTimePeriodSet As Boolean

' indicates whether grids in regions are currently
' hidden. Note that a region's hasGrid property
' indicates whether it has a grid, not whether it
' is currently visible
Private mHideGrid As Boolean

Private mPointerMode As PointerModes
Private mPointerStyle As PointerStyles
Private mPointerIcon As IPictureDisp
Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mSuppressDrawingCount As Long
Private mPainted As Boolean

'Private mCurrentTool As ToolTypes

Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mAllowHorizontalMouseScrolling As Boolean
Private mAllowVerticalMouseScrolling As Boolean

Private mMouseScrollingInProgress As Boolean

Private mShowHorizontalScrollBar As Boolean
Private mShowToolbar As Boolean

Private mRegionHeightReductionFactor As Double

Private mReferenceTime As Date

Private mAutoscroll As Boolean

Private mDefaultBarDisplayMode As BarDisplayModes

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Dim failpoint As Long
On Error GoTo Err

Set gBlankMouseIcon = BlankPicture.Picture

Set mController = New ChartController
mController.Chart = Me
initialise
createXAxisRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "UserControl_Initialize" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
mController.fireKeyDown KeyCode, Shift
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
mController.fireKeyPress KeyAscii
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
mController.fireKeyUp KeyCode, Shift
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

AllowHorizontalMouseScrolling = PropBag.ReadProperty(PropNameAllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling)
If Err.Number <> 0 Then
    AllowHorizontalMouseScrolling = PropDfltAllowHorizontalMouseScrolling
    Err.clear
End If

AllowVerticalMouseScrolling = PropBag.ReadProperty(PropNameAllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling)
If Err.Number <> 0 Then
    AllowVerticalMouseScrolling = PropDfltAllowVerticalMouseScrolling
    Err.clear
End If

Autoscroll = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
If Err.Number <> 0 Then
    Autoscroll = PropDfltAutoscroll
    Err.clear
End If

UserControl.BackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    UserControl.BackColor = PropDfltChartBackColor
    Err.clear
End If

DefaultBarDisplayMode = PropBag.ReadProperty(PropNameDefaultBarDisplayMode, PropDfltDefaultBarDisplayMode)
If Err.Number <> 0 Then
    DefaultBarDisplayMode = PropDfltDefaultBarDisplayMode
    Err.clear
End If

mXAxisRegion.BackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    mXAxisRegion.BackColor = PropDfltChartBackColor
    Err.clear
End If

mDefaultRegionStyle.Autoscale = PropBag.ReadProperty(PropNameDefaultRegionAutoscale, PropDfltDefaultRegionAutoscale)
If Err.Number <> 0 Then
    RegionDefaultAutoscale = PropDfltDefaultRegionAutoscale
    Err.clear
End If

RegionDefaultBackColor = PropBag.ReadProperty(PropNameDefaultRegionBackColor, PropDfltDefaultRegionBackColor)
If Err.Number <> 0 Then
    RegionDefaultBackColor = PropDfltDefaultRegionBackColor
    Err.clear
End If

RegionDefaultGridColor = PropBag.ReadProperty(PropNameDefaultRegionGridColor, PropDfltDefaultRegionGridColor)
If Err.Number <> 0 Then
    RegionDefaultGridColor = PropDfltDefaultRegionGridColor
    Err.clear
End If

RegionDefaultGridlineSpacingY = PropBag.ReadProperty(PropNameDefaultRegionGridlineSpacingY, PropDfltDefaultRegionGridlineSpacingY)
If Err.Number <> 0 Then
    RegionDefaultGridlineSpacingY = PropDfltDefaultRegionGridlineSpacingY
    Err.clear
End If

RegionDefaultGridTextColor = PropBag.ReadProperty(PropNameDefaultRegionGridTextColor, PropDfltDefaultRegionGridTextColor)
If Err.Number <> 0 Then
    RegionDefaultGridTextColor = PropDfltDefaultRegionGridTextColor
    Err.clear
End If

RegionDefaultHasGrid = PropBag.ReadProperty(PropNameDefaultRegionHasGrid, PropDfltDefaultRegionHasGrid)
If Err.Number <> 0 Then
    RegionDefaultHasGrid = PropDfltDefaultRegionHasGrid
    Err.clear
End If

RegionDefaultHasGridText = PropBag.ReadProperty(PropNameDefaultRegionHasGridtext, PropDfltDefaultRegionHasGridtext)
If Err.Number <> 0 Then
    RegionDefaultHasGridText = PropDfltDefaultRegionHasGridtext
    Err.clear
End If

RegionDefaultIntegerYScale = PropBag.ReadProperty(PropNameDefaultRegionIntegerYScale, PropDfltDefaultRegionIntegerYScale)
If Err.Number <> 0 Then
    RegionDefaultIntegerYScale = PropDfltDefaultRegionIntegerYScale
    Err.clear
End If

RegionDefaultMinimumHeight = PropBag.ReadProperty(PropNameDefaultRegionMinimumHeight, PropDfltDefaultRegionMinimumHeight)
If Err.Number <> 0 Then
    RegionDefaultMinimumHeight = PropDfltDefaultRegionMinimumHeight
    Err.clear
End If

RegionDefaultYScaleQuantum = PropBag.ReadProperty(PropNameDefaultRegionYScaleQuantum, PropDfltDefaultRegionYScaleQuantum)
If Err.Number <> 0 Then
    RegionDefaultYScaleQuantum = PropDfltDefaultRegionYScaleQuantum
    Err.clear
End If

'setPeriodParameters PropBag.ReadProperty(PropNamePeriodLength, PropDfltPeriodLength), _
'                    PropBag.ReadProperty(PropNamePeriodUnits, PropDfltPeriodUnits)
'If Err.Number <> 0 Then
'    setPeriodParameters PropDfltPeriodLength, PropDfltPeriodUnits
'    Err.clear
'End If

PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
If Err.Number <> 0 Then
    PointerCrosshairsColor = PropDfltPointerCrosshairsColor
    Err.clear
End If

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
If Err.Number <> 0 Then
    PointerDiscColor = PropDfltPointerDiscColor
    Err.clear
End If

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
If Err.Number <> 0 Then
    PointerStyle = PropDfltPointerStyle
    Err.clear
End If

ShowHorizontalScrollBar = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
If Err.Number <> 0 Then
    ShowHorizontalScrollBar = PropDfltShowHorizontalScrollBar
    Err.clear
End If

ShowToolbar = PropBag.ReadProperty(PropNameShowToolbar, PropDfltShowToolbar)
If Err.Number <> 0 Then
    ShowToolbar = PropDfltShowToolbar
    Err.clear
End If

TwipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)
If Err.Number <> 0 Then
    TwipsPerBar = PropDfltTwipsPerBar
    Err.clear
End If

Set mVerticalGridTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameVerticalGridSpacing, PropDfltVerticalGridSpacing), _
                        PropBag.ReadProperty(PropNameVerticalGridUnits, PropDfltVerticalGridUnits))
If Err.Number <> 0 Then
    Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
    Err.clear
End If

YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
If Err.Number <> 0 Then
    YAxisWidthCm = PropDfltYAxisWidthCm
    Err.clear
End If

End Sub

Private Sub UserControl_Resize()
Static resizeCount As Long

Dim failpoint As Long
On Error GoTo Err

'gLogger.Log LogLevelDetail, "ChartSkil: UserControl_Resize: enter"
resizeCount = resizeCount + 1

Resize (UserControl.width <> mPrevWidth), (UserControl.height <> mPrevHeight)
mPrevHeight = UserControl.height
mPrevWidth = UserControl.width

'gLogger.Log LogLevelDetail, "ChartSkil: UserControl_Resize: exit"

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "UserControl_Resize" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_Terminate()
Debug.Print "ChartSkil Usercontrol terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameAllowHorizontalMouseScrolling, AllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling
PropBag.WriteProperty PropNameAllowVerticalMouseScrolling, AllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling
PropBag.WriteProperty PropNameAutoscroll, Autoscroll, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, UserControl.BackColor, PropDfltChartBackColor
PropBag.WriteProperty PropNameDefaultBarDisplayMode, mDefaultBarDisplayMode, PropDfltDefaultBarDisplayMode
PropBag.WriteProperty PropNameDefaultRegionAutoscale, mDefaultRegionStyle.Autoscale, PropDfltDefaultRegionAutoscale
PropBag.WriteProperty PropNameDefaultRegionBackColor, mDefaultRegionStyle.BackColor, PropDfltDefaultRegionBackColor
PropBag.WriteProperty PropNameDefaultRegionGridColor, mDefaultRegionStyle.GridColor, PropDfltDefaultRegionGridColor
PropBag.WriteProperty PropNameDefaultRegionGridlineSpacingY, mDefaultRegionStyle.GridlineSpacingY, PropDfltDefaultRegionGridlineSpacingY
PropBag.WriteProperty PropNameDefaultRegionGridTextColor, mDefaultRegionStyle.GridTextColor, PropDfltDefaultRegionGridTextColor
PropBag.WriteProperty PropNameDefaultRegionHasGrid, mDefaultRegionStyle.HasGrid, PropDfltDefaultRegionHasGrid
PropBag.WriteProperty PropNameDefaultRegionHasGridtext, mDefaultRegionStyle.HasGridText, PropDfltDefaultRegionHasGridtext
PropBag.WriteProperty PropNameDefaultRegionIntegerYScale, mDefaultRegionStyle.IntegerYScale, PropDfltDefaultRegionIntegerYScale
PropBag.WriteProperty PropNameDefaultRegionMinimumHeight, mDefaultRegionStyle.MinimumHeight, PropDfltDefaultRegionMinimumHeight
PropBag.WriteProperty PropNameDefaultRegionYScaleQuantum, mDefaultRegionStyle.YScaleQuantum, PropDfltDefaultRegionYScaleQuantum
PropBag.WriteProperty PropNamePeriodLength, mBarTimePeriod.length, PropDfltPeriodLength
PropBag.WriteProperty PropNamePeriodUnits, mBarTimePeriod.units, PropDfltPeriodUnits
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNamePointerStyle, mPointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNameShowHorizontalScrollBar, ShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameShowToolbar, ShowToolbar, PropDfltShowToolbar
PropBag.WriteProperty PropNameTwipsPerBar, TwipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameVerticalGridSpacing, mVerticalGridTimePeriod.length, PropDfltVerticalGridSpacing
PropBag.WriteProperty PropNameVerticalGridUnits, mVerticalGridTimePeriod.units, PropDfltVerticalGridUnits
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

Set Region = mRegions(2 * index - 1).Region
Region.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_Click" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index - 1).Region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim Region As ChartRegion

Dim failpoint As Long

On Error GoTo Err

Set Region = mRegions(2 * index - 1).Region


If CBool(button And MouseButtonConstants.vbLeftButton) Then mMouseScrollingInProgress = True

' we notify the region selection first so that the application has a chance to
' turn off scrolling and snapping before getting the MouseDown event
RaiseEvent RegionSelected(Region)
mController.fireRegionSelected Region

If (mPointerMode = PointerModeDefault And _
        ((Region.SnapCursorToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.SnapCursorToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then y = YScaleQuantum * Int((y + YScaleQuantum / 10000) / YScaleQuantum)
End If

If mPointerMode = PointerModeDefault And _
    (mAllowHorizontalMouseScrolling Or mAllowVerticalMouseScrolling) _
Then
    mLeftDragStartPosnX = Int(x)
    mLeftDragStartPosnY = y
End If

Region.MouseDown button, Shift, Round(x), y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

Dim failpoint As Long
On Error GoTo Err

If CBool(button And MouseButtonConstants.vbLeftButton) Then
    If mPointerMode = PointerModeDefault And _
        (mAllowHorizontalMouseScrolling Or mAllowVerticalMouseScrolling) And _
        mMouseScrollingInProgress _
    Then
        mouseScroll index, button, Shift, x, y
    Else
        mMouseScrollingInProgress = False
        mouseMove index, button, Shift, x, y
    End If
Else
    mouseMove index, button, Shift, x, y
End If

mRegions(2 * index - 1).Region.mouseMove button, Shift, Round(x), y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err

mMouseScrollingInProgress = False

Set Region = mRegions(2 * index - 1).Region

If (mPointerMode = PointerModeDefault And _
        ((Region.SnapCursorToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.SnapCursorToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then y = YScaleQuantum * Int(y / YScaleQuantum)
End If

Region.MouseUp button, Shift, Round(x), y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseUp" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
Dim failpoint As Long
On Error GoTo Err

LastVisiblePeriod = Round((CLng(HScroll.value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.currentPeriodNumber + ChartWidth - 1))

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "HScroll_Change" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' RegionDividerPicture Event Handlers
'================================================================================

Private Sub RegionDividerPicture_MouseDown( _
                            index As Integer, _
                            button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If CBool(button And MouseButtonConstants.vbLeftButton) Then
    mLeftDragStartPosnX = Int(x)
    mLeftDragStartPosnY = y
    mUserResizingRegions = True
End If
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim vertChange As Long
Dim currRegion As Long
Dim newHeight As Long
Dim prevPercentHeight As Double
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If Not CBool(button And MouseButtonConstants.vbLeftButton) Then Exit Sub
If y = mLeftDragStartPosnY Then Exit Sub

' we resize the next region below the divider that has not
' been removed
For i = 2 * index + 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        currRegion = i
        Exit For
    End If
Next

vertChange = mLeftDragStartPosnY - y
newHeight = mRegions(currRegion).actualHeight + vertChange
If newHeight < 0 Then newHeight = 0

' the region table indicates the requested percentage used by each region
' and the actual height allocation. We need to work out the new percentage
' for the region to be resized.

'prevPercentHeight = mRegions(currRegion).region.percentheight

prevPercentHeight = mRegions(currRegion).percentheight
If Not mRegions(currRegion).useAvailableSpace Then
    'mRegions(currRegion).region.percentheight = prevPercentHeight * newHeight / mRegions(currRegion).actualHeight
    mRegions(currRegion).percentheight = 100 * newHeight / calcAvailableHeight
    'mRegions(currRegion).percentheight = prevPercentHeight * newHeight / mRegions(currRegion).actualHeight
Else
    ' this is a 'use available space' region that's being resized. Now change
    ' it to use a specific percentage
    mRegions(currRegion).Region.percentheight = 100 * newHeight / calcAvailableHeight
    mRegions(currRegion).percentheight = mRegions(currRegion).Region.percentheight
End If

If sizeRegions Then
    'paintAll
Else
    ' the regions couldn't be resized so reset the region's percent height
    mRegions(currRegion).percentheight = prevPercentHeight
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
mUserResizingRegions = False

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseUp" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Toolbar1 Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal button As MSComctlLib.button)

Dim failpoint As Long
On Error GoTo Err

Select Case button.Key
Case ToolbarCommandAutoScroll
    mAutoscroll = Not mAutoscroll
Case ToolbarCommandShowCrosshair
    PointerStyle = PointerCrosshairs
Case ToolbarCommandShowDiscCursor
    PointerStyle = PointerDisc
Case ToolbarCommandReduceSpacing
    If TwipsPerBar >= 50 Then
        TwipsPerBar = TwipsPerBar - 25
    End If
    If TwipsPerBar < 50 Then
        button.Enabled = False
    End If
Case ToolbarCommandIncreaseSpacing
    TwipsPerBar = TwipsPerBar + 25
    Toolbar1.Buttons("reducespacing").Enabled = True
Case ToolbarCommandScrollLeft
    ScrollX -(ChartWidth * 0.2)
Case ToolbarCommandScrollRight
    ScrollX ChartWidth * 0.2
Case ToolbarCommandScrollEnd
    LastVisiblePeriod = currentPeriodNumber
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Toolbar1_ButtonClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'================================================================================
' XAxisPicture Event Handlers
'================================================================================

Private Sub XAxisPicture_Click()
Dim failpoint As Long
On Error GoTo Err

mRegions(0).Region.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_Click" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_DblClick()
Dim failpoint As Long
On Error GoTo Err

mRegions(0).Region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseDown(button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).Region.MouseDown button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseMove(button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).Region.mouseMove button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseUp(button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).Region.MouseUp button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseUp" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' YAxisPicture Event Handlers
'================================================================================

Private Sub YAxisPicture_Click(index As Integer)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_Click" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseDown(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.MouseDown button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseMove(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.mouseMove button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseUp(index As Integer, button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.MouseUp button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseUp" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' mPeriods Event Handlers
'================================================================================

Private Sub mPeriods_PeriodAdded(ByVal period As period)
Dim i As Long
Dim Region As ChartRegion
Dim ev As CollectionChangeEvent

Dim failpoint As Long
On Error GoTo Err

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.addPeriod period.PeriodNumber, period.timestamp
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.addPeriod period.PeriodNumber, period.timestamp
If mSuppressDrawingCount = 0 Then setHorizontalScrollBar
setSession period.timestamp
If mAutoscroll Then ScrollX 1

Set ev.affectedItem = period
ev.changeType = CollItemAdded
Set ev.Source = mPeriods
RaiseEvent PeriodsChanged(ev)
mController.firePeriodsChanged ev

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mPeriods_PeriodAdded" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get AllowHorizontalMouseScrolling() As Boolean
Attribute AllowHorizontalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
AllowHorizontalMouseScrolling = mAllowHorizontalMouseScrolling
End Property

Public Property Let AllowHorizontalMouseScrolling(ByVal value As Boolean)
mAllowHorizontalMouseScrolling = value
PropertyChanged "allowHorizontalMouseScrolling"
End Property

Public Property Get AllowVerticalMouseScrolling() As Boolean
Attribute AllowVerticalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
AllowVerticalMouseScrolling = mAllowVerticalMouseScrolling
End Property

Public Property Let AllowVerticalMouseScrolling(ByVal value As Boolean)
mAllowVerticalMouseScrolling = value
PropertyChanged "allowVerticalMouseScrolling"
End Property

Public Property Get Autoscroll() As Boolean
Attribute Autoscroll.VB_ProcData.VB_Invoke_Property = ";Behavior"
Autoscroll = mAutoscroll
End Property

Public Property Let Autoscroll(ByVal value As Boolean)
mAutoscroll = value
End Property

Public Property Let BarTimePeriod( _
                ByVal value As TimePeriod)
If mBarTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "barTimePeriod", _
                                    "BarTimePeriod has already been set"
If value.length < 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "barTimePeriod", _
                                    "BarTimePeriod length cannot be negative"
                                    
Select Case value.units
Case TimePeriodNone
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case TimePeriodVolume
Case TimePeriodTickVolume
Case TimePeriodTickMovement
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            "ChartSkil" & "." & "Chart" & ":" & "setPeriodParameters", _
            "Invalid period unit - must be a member of the TimePeriodUnits enum"
End Select

Set mBarTimePeriod = value

mBarTimePeriodSet = True

If Not mVerticalGridTimePeriodSet Then calcVerticalGridParams
If mXAxisRegion Is Nothing Then createXAxisRegion
setRegionPeriodAndVerticalGridParameters

End Property

Public Property Get BarTimePeriod() As TimePeriod
Set BarTimePeriod = mBarTimePeriod
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
ChartBackColor = mDefaultRegionStyle.BackColor
End Property

Public Property Let ChartBackColor(ByVal val As OLE_COLOR)
Dim i As Long

If mDefaultRegionStyle.BackColor = val Then Exit Property

UserControl.BackColor = val

mDefaultRegionStyle.BackColor = val
mXAxisRegion.BackColor = val

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        mRegions(i).Region.BackColor = val
    End If
Next
paintAll
End Property

Public Property Get controller() As ChartController
Set controller = mController
End Property

Public Property Get ChartLeft() As Double
ChartLeft = mScaleLeft
End Property

Public Property Get ChartWidth() As Double
ChartWidth = YAxisPosition - mScaleLeft
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mPeriods.currentPeriodNumber
End Property

Public Property Get CurrentSessionEndTime() As Date
CurrentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
CurrentSessionStartTime = mCurrentSessionStartTime
End Property

'Public Property Get currentTool() As ToolTypes
'currentTool = mCurrentTool
'End Property

'Public Property Let currentTool(ByVal value As ToolTypes)
'Select Case value
'Case ToolPointer
'    mCurrentTool = value
'Case ToolLine
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolLineExtended
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolLineRay
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolLineHorizontal
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolLineVertical
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolFibonacciRetracement
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolFibonacciExtension
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolFibonacciCircle
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolFibonacciTime
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolRegressionChannel
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolRegressionEnvelope
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolText
'    mCurrentTool = ToolTypes.ToolPointer
'Case ToolPitchfork
'    mCurrentTool = ToolTypes.ToolPointer
'End Select
'End Property

Public Property Get DefaultBarDisplayMode() As BarDisplayModes
DefaultBarDisplayMode = mDefaultBarDisplayMode
End Property

Public Property Let DefaultBarDisplayMode( _
                ByVal value As BarDisplayModes)
mDefaultBarDisplayMode = value
End Property

Public Property Get DefaultBarStyle() As BarStyle
Set DefaultBarStyle = mDefaultBarStyle.clone
End Property

Public Property Let DefaultBarStyle( _
                ByVal value As BarStyle)
Set mDefaultRegionStyle = value.clone
End Property

Public Property Get DefaultDataPointStyle() As DataPointStyle
Set DefaultDataPointStyle = mDefaultDataPointStyle.clone
End Property

Public Property Let DefaultDataPointStyle( _
                ByVal value As DataPointStyle)
Set mDefaultDataPointStyle = value.clone
End Property

Public Property Get DefaultLineStyle() As linestyle
Set DefaultLineStyle = mDefaultLineStyle.clone
End Property

Public Property Let DefaultLineStyle( _
                ByVal value As linestyle)
Set mDefaultLineStyle = value.clone
End Property

Public Property Get DefaultRegionStyle() As ChartRegionStyle
Set DefaultRegionStyle = mDefaultRegionStyle.clone
End Property

Public Property Let DefaultRegionStyle( _
                ByVal value As ChartRegionStyle)
Set mDefaultRegionStyle = value.clone
End Property

Public Property Get DefaultTextStyle() As TextStyle
Set DefaultTextStyle = mDefaultTextStyle.clone
End Property

Public Property Let DefaultTextStyle(ByVal value As TextStyle)
Set mDefaultTextStyle = value.clone
End Property

Public Property Get DefaultYAxisStyle() As ChartRegionStyle
Set DefaultYAxisStyle = mDefaultYAxisStyle.clone
End Property

Public Property Let DefaultYAxisStyle(ByVal value As ChartRegionStyle)
Set mDefaultYAxisStyle = value.clone
End Property

Public Property Get FirstVisiblePeriod() As Long
FirstVisiblePeriod = mScaleLeft
End Property

Public Property Let FirstVisiblePeriod(ByVal value As Long)
ScrollX value - mScaleLeft + 1
End Property

Public Property Get LastVisiblePeriod() As Long
LastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let LastVisiblePeriod(ByVal value As Long)
ScrollX value - mYAxisPosition + 1
End Property

Public Property Get Periods() As Periods
Set Periods = mPeriods
End Property

Public Property Get PointerCrosshairsColor() As OLE_COLOR
Attribute PointerCrosshairsColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerCrosshairsColor = mPointerCrosshairsColor
End Property

Public Property Let PointerCrosshairsColor(ByVal value As OLE_COLOR)
Dim i As Long
Dim Region As ChartRegion
mPointerCrosshairsColor = value
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.PointerCrosshairsColor = value
    End If
Next
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Dim i As Long
Dim Region As ChartRegion
mPointerDiscColor = value
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.PointerDiscColor = value
    End If
Next
End Property

Public Property Get PointerIcon() As IPictureDisp
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)
Dim i As Long
Dim Region As ChartRegion

If value Is Nothing Then Exit Property
If value Is mPointerIcon Then Exit Property

Set mPointerIcon = value

If mPointerMode = PointerModeDefault Then
    If mPointerStyle = PointerCustom Then
        For i = 1 To mRegionsIndex Step 2
            If Not mRegions(i).Region Is Nothing Then
                Set Region = mRegions(i).Region
                Region.PointerIcon = value
                Region.PointerStyle = PointerCustom
            End If
        Next
    End If
End If
End Property

Public Property Get PointerStyle() As PointerStyles
Attribute PointerStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerStyle = mPointerStyle
End Property

Public Property Let PointerStyle(ByVal value As PointerStyles)
Dim i As Long
Dim Region As ChartRegion

If value = mPointerStyle Then Exit Property

mPointerStyle = value

If mPointerStyle = PointerCustom And mPointerIcon Is Nothing Then
    ' we'll notify the region when an icon is supplied
    Exit Property
End If

If mPointerMode = PointerModeDefault Then
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            Set Region = mRegions(i).Region
            If mPointerStyle = PointerCustom Then Region.PointerIcon = mPointerIcon
            Region.PointerStyle = mPointerStyle
        End If
    Next
End If
End Property

Public Property Get RegionDefaultAutoscale() As Boolean
Attribute RegionDefaultAutoscale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultAutoscale = mDefaultRegionStyle.Autoscale
End Property

Public Property Let RegionDefaultAutoscale(ByVal value As Boolean)
mDefaultRegionStyle.Autoscale = value
End Property

Public Property Get RegionDefaultBackColor() As OLE_COLOR
Attribute RegionDefaultBackColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultBackColor = mDefaultRegionStyle.BackColor
End Property

Public Property Let RegionDefaultBackColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.BackColor = val
End Property

Public Property Get RegionDefaultGridColor() As OLE_COLOR
Attribute RegionDefaultGridColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridColor = mDefaultRegionStyle.GridColor
End Property

Public Property Let RegionDefaultGridColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.GridColor = val
End Property

Public Property Get RegionDefaultGridlineSpacingY() As Double
Attribute RegionDefaultGridlineSpacingY.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridlineSpacingY = mDefaultRegionStyle.GridlineSpacingY
End Property

Public Property Let RegionDefaultGridlineSpacingY(ByVal value As Double)
mDefaultRegionStyle.GridlineSpacingY = value
End Property

Public Property Get RegionDefaultGridTextColor() As OLE_COLOR
Attribute RegionDefaultGridTextColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridTextColor = mDefaultRegionStyle.GridTextColor
End Property

Public Property Let RegionDefaultGridTextColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.GridTextColor = val
End Property

Public Property Get RegionDefaultHasGrid() As Boolean
Attribute RegionDefaultHasGrid.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGrid = mDefaultRegionStyle.HasGrid
End Property

Public Property Let RegionDefaultHasGrid(ByVal val As Boolean)
mDefaultRegionStyle.HasGrid = val
End Property

Public Property Get RegionDefaultHasGridText() As Boolean
Attribute RegionDefaultHasGridText.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGridText = mDefaultRegionStyle.HasGridText
End Property

Public Property Let RegionDefaultHasGridText(ByVal val As Boolean)
mDefaultRegionStyle.HasGridText = val
End Property

Public Property Get RegionDefaultIntegerYScale() As Boolean
Attribute RegionDefaultIntegerYScale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultIntegerYScale = mDefaultRegionStyle.IntegerYScale
End Property

Public Property Let RegionDefaultIntegerYScale(ByVal value As Boolean)
mDefaultRegionStyle.IntegerYScale = value
End Property

Public Property Get RegionDefaultMinimumHeight() As Double
Attribute RegionDefaultMinimumHeight.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultMinimumHeight = mDefaultRegionStyle.MinimumHeight
End Property

Public Property Let RegionDefaultMinimumHeight(ByVal value As Double)
mDefaultRegionStyle.MinimumHeight = value
End Property

Public Property Get RegionDefaultYScaleQuantum() As Double
Attribute RegionDefaultYScaleQuantum.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultYScaleQuantum = mDefaultRegionStyle.YScaleQuantum
End Property

Public Property Let RegionDefaultYScaleQuantum(ByVal value As Double)
mDefaultRegionStyle.YScaleQuantum = value
End Property

Public Property Get SessionEndTime() As Date
Attribute SessionEndTime.VB_ProcData.VB_Invoke_Property = ";Behavior"
SessionEndTime = mSessionEndTime
End Property

Public Property Let SessionEndTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionEndTime", _
                "Value must be a time only"
mSessionEndTime = val
End Property

Public Property Get SessionStartTime() As Date
Attribute SessionStartTime.VB_ProcData.VB_Invoke_Property = ";Behavior"
SessionStartTime = mSessionStartTime
End Property

Public Property Let SessionStartTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionStartTime", _
                "Value must be a time only"
mSessionStartTime = val
End Property

Public Property Get ShowHorizontalScrollBar() As Boolean
Attribute ShowHorizontalScrollBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
ShowHorizontalScrollBar = mShowHorizontalScrollBar
End Property

Public Property Let ShowHorizontalScrollBar(ByVal val As Boolean)
mShowHorizontalScrollBar = val
If mShowHorizontalScrollBar Then
    HScroll.Visible = True
Else
    HScroll.Visible = False
End If
Resize False, True
End Property

Public Property Get ShowToolbar() As Boolean
Attribute ShowToolbar.VB_ProcData.VB_Invoke_Property = ";Appearance"
ShowToolbar = mShowToolbar
End Property

Public Property Let ShowToolbar(ByVal val As Boolean)
mShowToolbar = val
If mShowToolbar Then
    Toolbar1.Visible = True
Else
    Toolbar1.Visible = False
End If
Resize False, True
End Property

Public Property Get SuppressDrawing() As Boolean
SuppressDrawing = (mSuppressDrawingCount > 0)
End Property

Public Property Let SuppressDrawing(ByVal val As Boolean)
Dim i As Long
Dim Region As ChartRegion
If val Then
    mSuppressDrawingCount = mSuppressDrawingCount + 1
Else
    If mSuppressDrawingCount > 0 Then
        mSuppressDrawingCount = mSuppressDrawingCount - 1
    End If
End If

If mSuppressDrawingCount = 0 Then
    Resize True, True
End If

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.SuppressDrawing = (mSuppressDrawingCount > 0)
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.SuppressDrawing = (mSuppressDrawingCount > 0)
End Property

Public Property Get TwipsPerBar() As Long
Attribute TwipsPerBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
TwipsPerBar = mTwipsPerBar
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
mTwipsPerBar = val
resizeX
setHorizontalScrollBar
'paintAll
End Property

Public Property Set VerticalGridTimePeriod( _
                ByVal value As TimePeriod)
If mVerticalGridTimePeriodSet Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                                    "verticalGridTimePeriod has already been set"

If value.length <= 0 Then Err.Raise ErrorCodes.ErrIllegalStateException, _
                                    "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                                    "verticalGridTimePeriod length must be >0"
Select Case value.units
Case TimePeriodSecond
Case TimePeriodMinute
Case TimePeriodHour
Case TimePeriodDay
Case TimePeriodWeek
Case TimePeriodMonth
Case TimePeriodYear
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil" & "." & "Chart" & ":" & "verticalGridTimePeriod", _
                "verticalGridTimePeriod Units must be a member of the TimePeriodUnits enum"
End Select

Set mVerticalGridTimePeriod = value
mVerticalGridTimePeriodSet = True

If mXAxisRegion Is Nothing Then createXAxisRegion
setRegionPeriodAndVerticalGridParameters

End Property

Public Property Get VerticalGridTimePeriod() As TimePeriod
Set VerticalGridTimePeriod = mVerticalGridTimePeriod
End Property

Public Property Get XAxisRegion() As ChartRegion
Set XAxisRegion = mXAxisRegion
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisWidthCm = mYAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
mYAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddChartRegion(ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double, _
                    Optional ByVal style As ChartRegionStyle, _
                    Optional ByVal yAxisStyle As ChartRegionStyle, _
                    Optional ByVal name As String) As ChartRegion
Dim ev As CollectionChangeEvent
Dim var As Variant
Dim p As period
Dim controlIndex As Long

'
' NB: percentHeight=100 means the region will use whatever space
' is available
'

Dim YAxisRegion As ChartRegion
Dim btn As button

If name <> "" Then
    If Not GetChartRegion(name) Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                "ChartSkil26.Chart::addChartRegion", _
                "Region " & name & " already exists"
    End If
End If

If style Is Nothing Then Set style = mDefaultRegionStyle
If yAxisStyle Is Nothing Then Set yAxisStyle = mDefaultYAxisStyle

mRegionsIndex = mRegionsIndex + 1
controlIndex = 1 + (mRegionsIndex - 1) / 2

Set AddChartRegion = New ChartRegion
AddChartRegion.name = name
AddChartRegion.controller = controller

Load ChartRegionPicture(controlIndex)
ChartRegionPicture(controlIndex).align = vbAlignNone
ChartRegionPicture(controlIndex).width = _
    UserControl.ScaleWidth * (mYAxisPosition - ChartLeft) / XAxisPicture.ScaleWidth
ChartRegionPicture(controlIndex).Visible = True

AddChartRegion.Canvas = createCanvas(ChartRegionPicture(controlIndex))

AddChartRegion.SuppressDrawing = (mSuppressDrawingCount > 0)
'addChartRegion.currentTool = mCurrentTool
AddChartRegion.minimumPercentHeight = minimumPercentHeight
AddChartRegion.percentheight = percentheight
AddChartRegion.PointerStyle = mPointerStyle
AddChartRegion.PointerCrosshairsColor = mPointerCrosshairsColor
AddChartRegion.PointerDiscColor = mPointerDiscColor
AddChartRegion.left = mScaleLeft
AddChartRegion.RegionNumber = mRegionsIndex
AddChartRegion.bottom = 0
AddChartRegion.top = 1
AddChartRegion.PeriodsInView mScaleLeft, mYAxisPosition - 1
AddChartRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
AddChartRegion.SessionStartTime = mSessionStartTime

AddChartRegion.DefaultBarStyle = mDefaultBarStyle
AddChartRegion.DefaultDataPointStyle = mDefaultDataPointStyle
AddChartRegion.DefaultLineStyle = mDefaultLineStyle
AddChartRegion.DefaultTextStyle = mDefaultTextStyle
AddChartRegion.style = style

If mHideGrid Then AddChartRegion.HideGrid

If mRegionsIndex = 1 Then
    For Each btn In Toolbar1.Buttons
        btn.Enabled = True
    Next
    Select Case mPointerStyle
    Case PointerNone
    
    Case PointerCrosshairs
        Toolbar1.Buttons("showcrosshair").value = tbrPressed
    Case PointerDisc
        Toolbar1.Buttons("showdisccursor").value = tbrPressed
    End Select
    AddChartRegion.Toolbar = Toolbar1
End If

Set mRegions(mRegionsIndex).Region = AddChartRegion
If percentheight <> 100 Then
    mRegions(mRegionsIndex).percentheight = mRegionHeightReductionFactor * percentheight
Else
    mRegions(mRegionsIndex).useAvailableSpace = True
End If

Load RegionDividerPicture(controlIndex)
RegionDividerPicture(controlIndex).Visible = True

mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then
    ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
End If

Set YAxisRegion = New ChartRegion
YAxisRegion.controller = controller

Load YAxisPicture(controlIndex)
YAxisPicture(controlIndex).align = vbAlignNone
YAxisPicture(controlIndex).left = ChartRegionPicture(controlIndex).width
YAxisPicture(controlIndex).width = UserControl.ScaleWidth - YAxisPicture(YAxisPicture.UBound).left
YAxisPicture(controlIndex).Visible = True

YAxisRegion.Canvas = createCanvas(YAxisPicture(controlIndex))

YAxisRegion.RegionNumber = mRegionsIndex
YAxisRegion.bottom = 0
YAxisRegion.top = 1
YAxisRegion.IsYAxisRegion = True
YAxisRegion.DefaultBarStyle = mDefaultBarStyle
YAxisRegion.DefaultDataPointStyle = mDefaultDataPointStyle
YAxisRegion.DefaultLineStyle = mDefaultLineStyle
YAxisRegion.DefaultTextStyle = mDefaultTextStyle
YAxisRegion.style = yAxisStyle
AddChartRegion.YAxisRegion = YAxisRegion

Set mRegions(mRegionsIndex).Region = YAxisRegion

mNumRegionsInUse = mNumRegionsInUse + 1

If sizeRegions Then
    Set ev.affectedItem = AddChartRegion
    ev.changeType = CollItemAdded
    RaiseEvent RegionsChanged(ev)
    mController.fireRegionsChanged ev
Else
    ' can't fit this all in! So remove the added region,
    Set AddChartRegion = Nothing
    Set mRegions(mRegionsIndex).Region = Nothing
    mRegions(mRegionsIndex).percentheight = 0
    mRegions(mRegionsIndex).actualHeight = 0
    mRegions(mRegionsIndex).useAvailableSpace = False
    Unload ChartRegionPicture(controlIndex)
    Unload RegionDividerPicture(mRegionsIndex)
    Unload YAxisPicture(controlIndex)
    mRegionsIndex = mRegionsIndex - 2
    mNumRegionsInUse = mNumRegionsInUse - 1
End If

End Function

Public Function addPeriod(ByVal timestamp As Date) As period
Set addPeriod = mPeriods.addPeriod(timestamp)
End Function

Public Function ClearChart()
Dim i As Long
Dim controlIndex As Long

For i = 1 To mRegionsIndex Step 2
    controlIndex = 1 + (i - 1) / 2
    If Not mRegions(i).Region Is Nothing Then
        mRegions(i).Region.ClearRegion
        mRegions(i + 1).Region.ClearRegion
        ChartRegionPicture(controlIndex).Visible = False
        YAxisPicture(controlIndex).Visible = False
        If i <> mRegionsIndex Then _
                RegionDividerPicture(controlIndex).Visible = False
    End If
Next

Erase mRegions

If Not mXAxisRegion Is Nothing Then mXAxisRegion.ClearRegion
XAxisPicture.Cls
Set mXAxisRegion = Nothing
mPeriods.Finish
Set mPeriods = Nothing

For i = 1 To ChartRegionPicture.UBound
    Unload ChartRegionPicture(i)
Next

For i = 1 To YAxisPicture.UBound
    Unload YAxisPicture(i)
Next

For i = 1 To RegionDividerPicture.UBound
    Unload RegionDividerPicture(i)
Next

initialise
mYAxisPosition = 1
createXAxisRegion
resizeX
'Resize False

RaiseEvent ChartCleared
mController.fireChartCleared
Debug.Print "Chart cleared"
End Function

Public Sub DisplayGrid()
Dim i As Long
Dim Region As ChartRegion

If Not mHideGrid Then Exit Sub

mHideGrid = False
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.DisplayGrid
    End If
Next
End Sub

Public Function GetChartRegion(ByVal name As String) As ChartRegion
Dim i As Long

name = UCase$(name)
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        If UCase$(mRegions(i).Region.name) = name Then
            Set GetChartRegion = mRegions(i).Region
            Exit Function
        End If
    End If
Next
                    
End Function

Public Sub HideGrid()
Dim i As Long
Dim Region As ChartRegion

If mHideGrid Then Exit Sub

mHideGrid = True
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.HideGrid
    End If
Next
End Sub

Public Function IsGridHidden() As Boolean
IsGridHidden = mHideGrid
End Function

Public Function IsTimeInSession(ByVal timestamp As Date) As Boolean

If timestamp >= mCurrentSessionStartTime And _
    timestamp < mCurrentSessionEndTime _
Then
    IsTimeInSession = True
End If
End Function

Public Sub RemoveChartRegion( _
                    ByVal Region As ChartRegion)
Dim i As Long
Dim ev As CollectionChangeEvent

Dim failpoint As Long
On Error GoTo Err

If Region.IsXAxisRegion Or Region.IsYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "removeChartRegion", _
            "Cannot remove an axis region"
End If

For i = 1 To mRegionsIndex Step 2
    If Region Is mRegions(i).Region Then
        Region.ClearRegion
        Set mRegions(i).Region = Nothing
        mRegions(i + 1).Region.ClearRegion
        Set mRegions(i + 1).Region = Nothing
        RegionDividerPicture(1 + (i - 1) / 2).Visible = False
        Exit For
    End If
Next

mNumRegionsInUse = mNumRegionsInUse - 1

sizeRegions

ev.changeType = CollItemRemoved
Set ev.affectedItem = Region
mController.fireRegionsChanged ev

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveChartRegion" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub ScrollX(ByVal value As Long)
Dim Region As ChartRegion
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If value = 0 Then Exit Sub

If (LastVisiblePeriod + value) > _
        (mPeriods.currentPeriodNumber + ChartWidth - 1) Then
    value = mPeriods.currentPeriodNumber + ChartWidth - 1 - LastVisiblePeriod
ElseIf (LastVisiblePeriod + value) < 1 Then
    value = 1 - LastVisiblePeriod
End If

mYAxisPosition = mYAxisPosition + value
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.width * mScaleWidth) - _
            mScaleWidth
XAxisPicture.ScaleLeft = mScaleLeft

If mSuppressDrawingCount > 0 Then Exit Sub

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        If Not mRegions(i).Region Is Nothing Then
            Set Region = mRegions(i).Region
            Region.PeriodsInView mScaleLeft, mYAxisPosition - 1
        End If
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.PeriodsInView mScaleLeft, mScaleLeft + mScaleWidth
setHorizontalScrollBar

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ScrollX" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetPointerModeDefault()
Dim i As Long
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

mPointerMode = PointerModeDefault
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.SetPointerModeDefault
    End If
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPointerModeDefault" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetPointerModeTool( _
                Optional ByVal toolPointerStyle As PointerStyles = PointerTool, _
                Optional ByVal icon As IPictureDisp)
Dim i As Long
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

mPointerMode = PointerModeTool

Select Case toolPointerStyle
Case PointerNone
Case PointerCrosshairs
Case PointerDisc
Case PointerTool
Case PointerCustom
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "SetPointerModeTool", _
            "toolPointerStyle must be a member of the PointerStyles enum"
End Select
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.SetPointerModeTool toolPointerStyle, icon
    End If
Next

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPointerModeTool" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcAvailableHeight() As Long
calcAvailableHeight = XAxisPicture.top - _
                    mNumRegionsInUse * RegionDividerPicture(0).height - _
                    IIf(mShowToolbar, Toolbar1.height, 0)
If calcAvailableHeight < 0 Then calcAvailableHeight = 0
End Function

Private Sub CalcSessionTimes(ByVal timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date)
Dim i As Long

i = -1
Do
    i = i + 1
Loop Until calcSessionTimesHelper(timestamp + i, SessionStartTime, SessionEndTime)
End Sub

Friend Function calcSessionTimesHelper(ByVal timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date) As Boolean
Dim referenceDate As Date
Dim referenceTime As Date
Dim weekday As Long

referenceDate = DateValue(timestamp)
referenceTime = TimeValue(timestamp)

If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If referenceTime < mSessionEndTime Then
        SessionStartTime = referenceDate + mSessionStartTime
        SessionEndTime = referenceDate + mSessionEndTime
    Else
        SessionStartTime = referenceDate + 1 + mSessionStartTime
        SessionEndTime = referenceDate + 1 + mSessionEndTime
    End If
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session spans midnight
    If referenceTime >= mSessionEndTime Then
        SessionStartTime = referenceDate + mSessionStartTime
        SessionEndTime = referenceDate + 1 + mSessionEndTime
    Else
        SessionStartTime = referenceDate - 1 + mSessionStartTime
        SessionEndTime = referenceDate + mSessionEndTime
    End If
Else
    ' this instrument trades 24hrs, or the contract service provider doesn't know
    ' the session start and end times
    SessionStartTime = referenceDate
    SessionEndTime = referenceDate + 1
End If

weekday = DatePart("w", SessionStartTime)
If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session DOES span midnight
    If weekday <> vbFriday And weekday <> vbSaturday Then calcSessionTimesHelper = True
Else
    ' 24-hour session or no session times known
    If weekday <> vbSaturday And weekday <> vbSunday Then calcSessionTimesHelper = True
End If
End Function

Private Sub calcVerticalGridParams()

Select Case mBarTimePeriod.units
Case TimePeriodNone
    Set mVerticalGridTimePeriod = Nothing
Case TimePeriodSecond
    Select Case mBarTimePeriod.length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodSecond)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodSecond)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(20, TimePeriodSecond)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMinute)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(5, TimePeriodMinute)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodMinute
    Select Case mBarTimePeriod.length
    Case 1
        Set mVerticalGridTimePeriod = GetTimePeriod(15, TimePeriodMinute)
    Case 2
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 3
        Set mVerticalGridTimePeriod = GetTimePeriod(30, TimePeriodMinute)
    Case 4
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 5
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 6
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodHour)
    Case 10
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 12
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 15
        Set mVerticalGridTimePeriod = GetTimePeriod(2, TimePeriodHour)
    Case 20
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case 30
        Set mVerticalGridTimePeriod = GetTimePeriod(4, TimePeriodHour)
    Case Else
        Set mVerticalGridTimePeriod = Nothing
    End Select
Case TimePeriodHour
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodDay)
Case TimePeriodDay
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodWeek)
Case TimePeriodWeek
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodMonth)
Case TimePeriodMonth
        Set mVerticalGridTimePeriod = GetTimePeriod(1, TimePeriodYear)
Case TimePeriodYear
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodYear)
Case TimePeriodVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodVolume)
Case TimePeriodTickVolume
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickVolume)
Case TimePeriodTickMovement
        Set mVerticalGridTimePeriod = GetTimePeriod(10, TimePeriodTickMovement)
End Select
  
End Sub

Private Function createCanvas( _
                ByVal surface As PictureBox) As Canvas
Set createCanvas = New Canvas
createCanvas.surface = surface
End Function

Private Sub createXAxisRegion()
Dim aFont As StdFont
Dim style As ChartRegionStyle

Set mXAxisRegion = New ChartRegion
mXAxisRegion.IsXAxisRegion = True

Set mRegions(0).Region = mXAxisRegion

mXAxisRegion.controller = controller
mXAxisRegion.Canvas = createCanvas(XAxisPicture)
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
mXAxisRegion.bottom = 0
mXAxisRegion.top = 1
mXAxisRegion.SessionStartTime = mSessionStartTime

mXAxisRegion.DefaultBarStyle = mDefaultBarStyle
mXAxisRegion.DefaultDataPointStyle = mDefaultDataPointStyle
mXAxisRegion.DefaultLineStyle = mDefaultLineStyle
mXAxisRegion.DefaultTextStyle = mDefaultTextStyle

Set style = mDefaultRegionStyle.clone
style.HasGrid = False
style.HasGridText = True
mXAxisRegion.style = style

Set mXCursorText = mXAxisRegion.AddText(LayerNumbers.LayerPointer)
mXCursorText.align = AlignTopCentre
mXCursorText.Color = vbWhite Xor mDefaultRegionStyle.BackColor
mXCursorText.box = True
mXCursorText.boxFillColor = vbWhite
mXCursorText.boxStyle = LineSolid
mXCursorText.boxColor = vbBlack
Set aFont = New StdFont
aFont.name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
mXCursorText.font = aFont

End Sub

Private Sub displayXAxisLabel(ByVal x As Single, ByVal y As Single)
Dim thisPeriod As period
Dim PeriodNumber As Long
Dim prevPeriodNumber As Long
Dim prevPeriod As period

If mXAxisRegion Is Nothing Then createXAxisRegion

If Round(x) >= mYAxisPosition Then Exit Sub
If mPeriods.Count = 0 Then Exit Sub

On Error Resume Next
PeriodNumber = Round(x)
Set thisPeriod = mPeriods(PeriodNumber)
On Error GoTo 0
If thisPeriod Is Nothing Then
    mXCursorText.text = ""
    Exit Sub
End If

mXCursorText.position = mXAxisRegion.newPoint( _
                            PeriodNumber, _
                            0, _
                            CoordsLogical, _
                            CoordsCounterDistance)

Select Case mBarTimePeriod.units
Case TimePeriodNone, TimePeriodMinute, TimePeriodHour
    mXCursorText.text = FormatDateTime(thisPeriod.timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.timestamp, vbShortTime)
Case TimePeriodSecond, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    mXCursorText.text = FormatDateTime(thisPeriod.timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.timestamp, vbLongTime)
Case Else
    mXCursorText.text = FormatDateTime(thisPeriod.timestamp, vbShortDate)
End Select

End Sub

Private Sub initialise()
Static firstInitialisationDone As Boolean
Dim i As Long
Dim btn As button

For Each btn In Toolbar1.Buttons
    btn.value = tbrUnpressed
    btn.Enabled = False
Next

mPrevHeight = UserControl.height

ReDim mRegions(3) As RegionTableEntry
mRegionsIndex = 0
mNumRegionsInUse = 0
mRegionHeightReductionFactor = 1

Set mPeriods = New Periods
mPeriods.controller = controller

mBarTimePeriodSet = False

If Not firstInitialisationDone Then
    
    firstInitialisationDone = True
    
    ' these values are only set once when the control initialises
    ' if the chart is subsequently cleared, any values set by the
    ' application remain in force
    mAutoscroll = PropDfltAutoscroll
    Set mBarTimePeriod = GetTimePeriod(PropDfltPeriodLength, PropDfltPeriodUnits)
    mPointerCrosshairsColor = PropDfltPointerCrosshairsColor
    mPointerDiscColor = PropDfltPointerDiscColor
    mShowHorizontalScrollBar = PropDfltShowHorizontalScrollBar
    mShowToolbar = PropDfltShowToolbar
    'HScroll.height = HorizScrollBarHeight
    HScroll.Visible = mShowHorizontalScrollBar
    Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
    mVerticalGridTimePeriodSet = False
    
    Set mDefaultRegionStyle = gCreateChartRegionStyle
    
    Set mDefaultBarStyle = gCreateBarStyle
    
    Set mDefaultDataPointStyle = gCreateDataPointStyle
    
    Set mDefaultLineStyle = gCreateLineStyle
    
    Set mDefaultTextStyle = gCreateTextStyle
    
    Set mDefaultYAxisStyle = gCreateChartRegionStyle
    mDefaultYAxisStyle.HasGrid = False
    
    mTwipsPerBar = PropDfltTwipsPerBar
    mYAxisWidthCm = PropDfltYAxisWidthCm

    mAllowHorizontalMouseScrolling = PropDfltAllowHorizontalMouseScrolling
    mAllowVerticalMouseScrolling = PropDfltAllowVerticalMouseScrolling

End If

mYAxisPosition = 1
mScaleWidth = CSng(XAxisPicture.width) / CSng(mTwipsPerBar) - 0.5!
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.width * mScaleWidth) - _
            mScaleWidth
mScaleHeight = -100
mScaleTop = 100

HScroll.value = 0
'resizeX


End Sub

Private Sub mouseMove( _
                ByVal index As Long, _
                ByVal button As Long, _
                ByVal Shift As Long, _
                ByRef x As Single, _
                ByRef y As Single)
Dim i As Long
Dim Region As ChartRegion

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        If i = (2 * index - 1) Then
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & y
            If (mPointerMode = PointerModeDefault And _
                    ((Region.SnapCursorToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
                    (Not Region.SnapCursorToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
                (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
            Then
                Dim YScaleQuantum As Double
                YScaleQuantum = Region.YScaleQuantum
                If YScaleQuantum <> 0 Then y = YScaleQuantum * Int((y + YScaleQuantum / 10000) / YScaleQuantum)
            End If
            Region.DrawCursor button, Shift, x, y
            
        Else
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & MinusInfinitySingle
            Region.DrawCursor button, Shift, x, MinusInfinitySingle
        End If
    End If
Next
displayXAxisLabel Round(x), 100
End Sub

Private Sub mouseScroll( _
                ByVal index As Long, _
                ByVal button As Long, _
                ByVal Shift As Long, _
                ByRef x As Single, _
                ByRef y As Single)

If mAllowHorizontalMouseScrolling Then
    ' the chart needs to be scrolled so that current mouse position
    ' is the value contained in mLeftDragStartPosnX
    If mLeftDragStartPosnX <> Int(x) Then
        If (LastVisiblePeriod + mLeftDragStartPosnX - Int(x)) <= _
                (mPeriods.currentPeriodNumber + ChartWidth - 1) And _
            (LastVisiblePeriod + mLeftDragStartPosnX - Int(x)) >= 1 _
        Then
            ScrollX mLeftDragStartPosnX - Int(x)
        End If
    End If
End If
If mAllowVerticalMouseScrolling Then
    If mLeftDragStartPosnY <> y Then
        With mRegions(2 * index - 1).Region
            If Not .Autoscale Then
                .ScrollVertical mLeftDragStartPosnY - y
            End If
        End With
    End If
End If
End Sub

Private Sub paintAll()
Dim Region As ChartRegion
Dim i As Long

If mSuppressDrawingCount > 0 Then Exit Sub

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.PaintRegion
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.PaintRegion

End Sub

Private Sub Resize( _
    ByVal resizeWidth As Boolean, _
    ByVal resizeHeight As Boolean)
Dim failpoint As Long
On Error GoTo Err

failpoint = 100

gLogger.Log LogLevelDetail, "ChartSkil: Resize: enter"

If resizeWidth Then
    HScroll.width = UserControl.width
    XAxisPicture.width = UserControl.width
    Toolbar1.width = UserControl.width
    resizeX
End If

failpoint = 200

If resizeHeight Then
    HScroll.top = UserControl.height - IIf(mShowHorizontalScrollBar, HScroll.height, 0)
    XAxisPicture.top = HScroll.top - XAxisPicture.height
    sizeRegions
End If
'paintAll

gLogger.Log LogLevelDetail, "ChartSkil: Resize: exit"

Exit Sub

Err:
gLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "Resize" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "Resize" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub resizeX()
Dim newScaleWidth As Single
Dim i As Long
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err


failpoint = 100

If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":resizeX Enter"


failpoint = 200

newScaleWidth = CSng(XAxisPicture.width) / CSng(mTwipsPerBar) - 0.5!
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.width * newScaleWidth) - _
            newScaleWidth


failpoint = 300

mScaleWidth = newScaleWidth


failpoint = 400

For i = 0 To ChartRegionPicture.UBound
    If (UserControl.width - YAxisPicture(i).width) > 0 Then
        YAxisPicture(i).left = UserControl.width - YAxisPicture(i).width
        ChartRegionPicture(i).width = YAxisPicture(i).left
    End If
Next


failpoint = 500

For i = 0 To RegionDividerPicture.UBound
    RegionDividerPicture(i).width = UserControl.width
Next


failpoint = 600

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.PeriodsInView mScaleLeft, mYAxisPosition - 1
    End If
Next

failpoint = 700

If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.PeriodsInView mScaleLeft, mScaleLeft + mScaleWidth
End If


failpoint = 800

setHorizontalScrollBar

If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":resizeX Exit"

Exit Sub

Err:
gLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "resizeX" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "resizeX" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub setHorizontalScrollBar()
Dim failpoint As Long
Dim hscrollVal As Integer
On Error GoTo Err

If mPeriods.currentPeriodNumber + ChartWidth - 1 > 32767 Then

    failpoint = 100

    HScroll.Max = 32767
ElseIf mPeriods.currentPeriodNumber + ChartWidth - 1 < 1 Then

    failpoint = 200

    HScroll.Max = 1
Else

    failpoint = 300
    
    HScroll.Max = mPeriods.currentPeriodNumber + ChartWidth - 1
End If
HScroll.Min = 0


failpoint = 400

' NB the following calculation has to be done using doubles as for very large charts it can cause an overflow using integers
hscrollVal = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.currentPeriodNumber + ChartWidth - 1)))
If hscrollVal > HScroll.Max Then
    HScroll.value = HScroll.Max
ElseIf hscrollVal < HScroll.Min Then
    HScroll.value = HScroll.Min
Else
    HScroll.value = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.currentPeriodNumber + ChartWidth - 1)))
End If

failpoint = 500

HScroll.SmallChange = 1
If (ChartWidth - 1) < 1 Then
    HScroll.LargeChange = 1
Else
    HScroll.LargeChange = ChartWidth - 1
End If

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = Err.Source
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "setHorizontalScrollBar" & "." & failpoint & _
                            IIf(errSource <> "", vbCrLf & errSource, "") & vbCrLf & _
                            errDescription
Err.Raise errNumber, _
        ProjectName & "." & ModuleName & ":" & "mTimer_TimerExpired" & "." & failpoint & _
        IIf(errSource <> "", vbCrLf & errSource, ""), _
        errDescription

End Sub

Private Sub setSession( _
                ByVal timestamp As Date)
If timestamp >= mCurrentSessionEndTime Or _
    timestamp < mReferenceTime _
Then
    mReferenceTime = timestamp
    CalcSessionTimes timestamp, mCurrentSessionStartTime, mCurrentSessionEndTime
End If
End Sub

Private Sub setRegionPeriodAndVerticalGridParameters()
Dim i As Long
Dim Region As ChartRegion
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.VerticalGridTimePeriod = mVerticalGridTimePeriod
    End If
Next
End Sub

Private Function sizeRegions() As Boolean
'
' NB: percentHeight=100 means the region will use whatever space
' is available
'
Dim i As Long
Dim top As Long
Dim aRegion As ChartRegion
Dim numAvailableSpaceRegions As Long
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim availableHeight As Long     ' the space available for the region picture boxes
                                ' excluding the divider pictures
Dim numRegionsSized As Long
Dim heightReductionFactor As Double
Dim failpoint As Long
On Error GoTo Err

If gLogger.isLoggable(LogLevelHighDetail) Then gLogger.Log LogLevelHighDetail, ProjectName & "." & ModuleName & ":sizeRegions Enter"


failpoint = 100

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
'        mRegions(i).percentheight = aRegion.percentheight
        If Not mRegions(i).useAvailableSpace Then
            availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
        Else
            If aRegion.minimumPercentHeight <> 0 Then
                availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
            End If
            numAvailableSpaceRegions = numAvailableSpaceRegions + 1
        End If
    End If
Next

If availableSpacePercent < 0 And mUserResizingRegions Then
    sizeRegions = False
    Exit Function
End If


failpoint = 200

heightReductionFactor = 1
Do While availableSpacePercent < 0
    availableSpacePercent = 100
    nonFixedAvailableSpacePercent = 100
    mRegionHeightReductionFactor = mRegionHeightReductionFactor * 0.95
    heightReductionFactor = heightReductionFactor * 0.95
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            Set aRegion = mRegions(i).Region
            If Not mRegions(i).useAvailableSpace Then
                If aRegion.minimumPercentHeight <> 0 Then
                    If mRegions(i).percentheight * mRegionHeightReductionFactor >= _
                        aRegion.minimumPercentHeight _
                    Then
                        mRegions(i).percentheight = mRegions(i).percentheight * mRegionHeightReductionFactor
                    Else
                        mRegions(i).percentheight = aRegion.minimumPercentHeight
                    End If
                    totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
                Else
                    mRegions(i).percentheight = mRegions(i).percentheight * mRegionHeightReductionFactor
                End If
                availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
                nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
            Else
                If aRegion.minimumPercentHeight <> 0 Then
                    availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
                    totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
                End If
            End If
        End If
    Next
    If totalMinimumPercents > 100 Then
        ' can't possibly fit this all in!
        sizeRegions = False
        If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":sizeRegions Exit"
        Exit Function
    End If
Loop


failpoint = 300

If numAvailableSpaceRegions = 0 Then
    ' we must adjust the percentages on the other regions so they
    ' total 100.
    For i = 1 To mRegionsIndex Step 2
        mRegions(i).percentheight = 100 * mRegions(i).percentheight / (100 - nonFixedAvailableSpacePercent)
    Next
End If

' calculate the actual available height to put these regions in
availableHeight = calcAvailableHeight

' first set heights for fixed height regions

failpoint = 400

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).useAvailableSpace Then
        mRegions(i).actualHeight = mRegions(i).percentheight * availableHeight / 100
        Debug.Assert mRegions(i).actualHeight >= 0
    End If
Next


failpoint = 500

' now set heights for 'available space' regions with a minimum height
' that needs to be respected
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If mRegions(i).useAvailableSpace Then
            mRegions(i).actualHeight = 0
            If aRegion.minimumPercentHeight <> 0 Then
                If (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) < aRegion.minimumPercentHeight Then
                    mRegions(i).actualHeight = aRegion.minimumPercentHeight * availableHeight / 100
                    Debug.Assert mRegions(i).actualHeight >= 0
                    nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.minimumPercentHeight
                    numAvailableSpaceRegions = numAvailableSpaceRegions - 1
                End If
            End If
        End If
    End If
Next


failpoint = 600

' finally set heights for all other 'available space' regions
For i = 1 To mRegionsIndex Step 2
    If mRegions(i).useAvailableSpace And _
        mRegions(i).actualHeight = 0 _
    Then
        mRegions(i).actualHeight = (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) * availableHeight / 100
        Debug.Assert mRegions(i).actualHeight >= 0
    End If
Next


failpoint = 700

' Now actually set the heights and positions for the picture boxes

top = IIf(mShowToolbar, Toolbar1.height, 0)

Dim controlIndex As Long
    
For i = 1 To mRegionsIndex Step 2
    controlIndex = 1 + (i - 1) / 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If Not SuppressDrawing Then
            ChartRegionPicture(controlIndex).height = mRegions(i).actualHeight
            YAxisPicture(controlIndex).height = mRegions(i).actualHeight
            ChartRegionPicture(controlIndex).top = top
            YAxisPicture(controlIndex).top = top
            aRegion.resizedY
        End If
        top = top + mRegions(i).actualHeight
        numRegionsSized = numRegionsSized + 1
        If Not SuppressDrawing Then
            RegionDividerPicture(controlIndex).top = top
        End If
        If numRegionsSized <> mNumRegionsInUse Then
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbSizeNS
        Else
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbDefault
        End If
        top = top + RegionDividerPicture(controlIndex).height
    Else
        If Not SuppressDrawing Then
            ChartRegionPicture(controlIndex).Visible = False
            YAxisPicture(controlIndex).Visible = False
            RegionDividerPicture(controlIndex).Visible = False
        End If
    End If
Next

sizeRegions = True

If gLogger.isLoggable(LogLevelHighDetail) Then gLogger.Log LogLevelHighDetail, ProjectName & "." & ModuleName & ":sizeRegions Exit"

Exit Function

Err:
gLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "sizeRegions" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "sizeRegions" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Private Sub zoom(ByRef rect As TRectangle)

End Sub

