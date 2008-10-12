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
            Picture         =   "ChartArea.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":159A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":19EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1E3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2290
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":26E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2F86
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":33D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":382A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":3C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":40CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4520
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
            Picture         =   "ChartArea.ctx":4972
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4DC4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5216
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5668
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5ABA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":635E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":67B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":6C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7054
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":74A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":78F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7D4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":819C
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":85EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8A40
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8E92
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
            Picture         =   "ChartArea.ctx":92E4
            Key             =   "showbars"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":95FE
            Key             =   "showcandlesticks"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9918
            Key             =   "showline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9C32
            Key             =   "showcrosshair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9F4C
            Key             =   "showdisccursor"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A266
            Key             =   "thinnerbars"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A580
            Key             =   "thickerbars"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A89A
            Key             =   "narrower"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":ACEC
            Key             =   "wider"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B006
            Key             =   "scaledown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B320
            Key             =   "scaleup"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B63A
            Key             =   "scrolldown"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B954
            Key             =   "scrollup"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":BC6E
            Key             =   "scrollleft"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":BF88
            Key             =   "scrollright"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C2A2
            Key             =   "scrollend"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C5BC
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
            Picture         =   "ChartArea.ctx":C8D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CBF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CF0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D224
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D53E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D858
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DB72
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DE8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E2DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E730
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":EA4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":ED64
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F07E
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F398
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F6B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F9CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":FCE6
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

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

Private Type RegionTableEntry
    region              As ChartRegion
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
Private Const PropNameDefaultRegionPointerStyle         As String = "DefaultRegionPointerStyle"
Private Const PropNameDefaultRegionYScaleQuantum        As String = "DefaultRegionYScaleQuantum"
Private Const PropNamePeriodLength                      As String = "PeriodLength"
Private Const PropNamePeriodUnits                       As String = "PeriodUnits"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
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
Private Const PropDfltDefaultRegionPointerStyle         As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltDefaultRegionYScaleQuantum        As Double = 0#
Private Const PropDfltPeriodLength                      As Long = 5
Private Const PropDfltPeriodUnits                       As Long = TimePeriodMinute
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
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

Private mPointerStyle As PointerStyles
Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mSuppressDrawingCount As Long
Private mPainted As Boolean

Private mCurrentTool As ToolTypes

Private mLeftDragging As Boolean    ' set when the mouse is being dragged with
                                    ' the left button depressed
Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mAllowHorizontalMouseScrolling As Boolean
Private mAllowVerticalMouseScrolling As Boolean

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
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

allowHorizontalMouseScrolling = PropBag.ReadProperty(PropNameAllowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling)
If Err.Number <> 0 Then
    allowHorizontalMouseScrolling = PropDfltAllowHorizontalMouseScrolling
    Err.clear
End If

allowVerticalMouseScrolling = PropBag.ReadProperty(PropNameAllowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling)
If Err.Number <> 0 Then
    allowVerticalMouseScrolling = PropDfltAllowVerticalMouseScrolling
    Err.clear
End If

autoscroll = PropBag.ReadProperty(PropNameAutoscroll, PropDfltAutoscroll)
If Err.Number <> 0 Then
    autoscroll = PropDfltAutoscroll
    Err.clear
End If

UserControl.backColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    UserControl.backColor = PropDfltChartBackColor
    Err.clear
End If

defaultBarDisplayMode = PropBag.ReadProperty(PropNameDefaultBarDisplayMode, PropDfltDefaultBarDisplayMode)
If Err.Number <> 0 Then
    defaultBarDisplayMode = PropDfltDefaultBarDisplayMode
    Err.clear
End If

mXAxisRegion.regionBackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    mXAxisRegion.regionBackColor = PropDfltChartBackColor
    Err.clear
End If

mDefaultRegionStyle.autoscale = PropBag.ReadProperty(PropNameDefaultRegionAutoscale, PropDfltDefaultRegionAutoscale)
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

RegionDefaultPointerStyle = PropBag.ReadProperty(PropNameDefaultRegionPointerStyle, PropDfltDefaultRegionPointerStyle)
If Err.Number <> 0 Then
    RegionDefaultPointerStyle = PropDfltDefaultRegionPointerStyle
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

showHorizontalScrollBar = PropBag.ReadProperty(PropNameShowHorizontalScrollBar, PropDfltShowHorizontalScrollBar)
If Err.Number <> 0 Then
    showHorizontalScrollBar = PropDfltShowHorizontalScrollBar
    Err.clear
End If

showToolbar = PropBag.ReadProperty(PropNameShowToolbar, PropDfltShowToolbar)
If Err.Number <> 0 Then
    showToolbar = PropDfltShowToolbar
    Err.clear
End If

twipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)
If Err.Number <> 0 Then
    twipsPerBar = PropDfltTwipsPerBar
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
PropBag.WriteProperty PropNameAllowHorizontalMouseScrolling, allowHorizontalMouseScrolling, PropDfltAllowHorizontalMouseScrolling
PropBag.WriteProperty PropNameAllowVerticalMouseScrolling, allowVerticalMouseScrolling, PropDfltAllowVerticalMouseScrolling
PropBag.WriteProperty PropNameAutoscroll, autoscroll, PropDfltAutoscroll
PropBag.WriteProperty PropNameChartBackColor, UserControl.backColor, PropDfltChartBackColor
PropBag.WriteProperty PropNameDefaultBarDisplayMode, mDefaultBarDisplayMode, PropDfltDefaultBarDisplayMode
PropBag.WriteProperty PropNameDefaultRegionAutoscale, mDefaultRegionStyle.autoscale, PropDfltDefaultRegionAutoscale
PropBag.WriteProperty PropNameDefaultRegionBackColor, mDefaultRegionStyle.backColor, PropDfltDefaultRegionBackColor
PropBag.WriteProperty PropNameDefaultRegionGridColor, mDefaultRegionStyle.gridColor, PropDfltDefaultRegionGridColor
PropBag.WriteProperty PropNameDefaultRegionGridlineSpacingY, mDefaultRegionStyle.gridlineSpacingY, PropDfltDefaultRegionGridlineSpacingY
PropBag.WriteProperty PropNameDefaultRegionGridTextColor, mDefaultRegionStyle.gridTextColor, PropDfltDefaultRegionGridTextColor
PropBag.WriteProperty PropNameDefaultRegionHasGrid, mDefaultRegionStyle.hasGrid, PropDfltDefaultRegionHasGrid
PropBag.WriteProperty PropNameDefaultRegionHasGridtext, mDefaultRegionStyle.hasGridText, PropDfltDefaultRegionHasGridtext
PropBag.WriteProperty PropNameDefaultRegionIntegerYScale, mDefaultRegionStyle.integerYScale, PropDfltDefaultRegionIntegerYScale
PropBag.WriteProperty PropNameDefaultRegionMinimumHeight, mDefaultRegionStyle.minimumHeight, PropDfltDefaultRegionMinimumHeight
PropBag.WriteProperty PropNameDefaultRegionPointerStyle, mDefaultRegionStyle.pointerStyle, PropDfltDefaultRegionPointerStyle
PropBag.WriteProperty PropNameDefaultRegionYScaleQuantum, mDefaultRegionStyle.YScaleQuantum, PropDfltDefaultRegionYScaleQuantum
PropBag.WriteProperty PropNamePeriodLength, mBarTimePeriod.length, PropDfltPeriodLength
PropBag.WriteProperty PropNamePeriodUnits, mBarTimePeriod.units, PropDfltPeriodUnits
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNameShowHorizontalScrollBar, showHorizontalScrollBar, PropDfltShowHorizontalScrollBar
PropBag.WriteProperty PropNameShowToolbar, showToolbar, PropDfltShowToolbar
PropBag.WriteProperty PropNameTwipsPerBar, twipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameVerticalGridSpacing, mVerticalGridTimePeriod.length, PropDfltVerticalGridSpacing
PropBag.WriteProperty PropNameVerticalGridUnits, mVerticalGridTimePeriod.units, PropDfltVerticalGridUnits
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index - 1).region.Click

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

mRegions(2 * index - 1).region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim region As ChartRegion

Dim failpoint As Long

On Error GoTo Err

Set region = mRegions(2 * index - 1).region

If (region.snapCursorToTickBoundaries And Not CBool((Shift And vbCtrlMask))) Or _
    (Not region.snapCursorToTickBoundaries And CBool((Shift And vbCtrlMask))) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = region.YScaleQuantum
    If YScaleQuantum <> 0 Then y = YScaleQuantum * Int(y / YScaleQuantum)
End If

If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(x)
mLeftDragStartPosnY = y

region.MouseDown Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

Dim region As ChartRegion
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If mLeftDragging = True Then
    If mAllowHorizontalMouseScrolling Then
        ' the chart needs to be scrolled so that current mouse position
        ' is the value contained in mLeftDragStartPosnX
        If mLeftDragStartPosnX <> Int(x) Then
            If (lastVisiblePeriod + mLeftDragStartPosnX - Int(x)) <= _
                    (mPeriods.currentPeriodNumber + chartWidth - 1) And _
                (lastVisiblePeriod + mLeftDragStartPosnX - Int(x)) >= 1 _
            Then
                scrollX mLeftDragStartPosnX - Int(x)
            End If
        End If
    End If
    If mAllowVerticalMouseScrolling Then
        If mLeftDragStartPosnY <> y Then
            With mRegions(2 * index - 1).region
                If Not .autoscale Then
                    .scrollVertical mLeftDragStartPosnY - y
                End If
            End With
        End If
    End If
Else
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).region Is Nothing Then
            Set region = mRegions(i).region
            If i = (2 * index - 1) Then
                'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & y
                If (region.snapCursorToTickBoundaries And Not CBool((Shift And vbCtrlMask))) Or _
                    (Not region.snapCursorToTickBoundaries And CBool((Shift And vbCtrlMask))) _
                Then
                    Dim YScaleQuantum As Double
                    YScaleQuantum = region.YScaleQuantum
                    If YScaleQuantum <> 0 Then y = YScaleQuantum * Int((y + YScaleQuantum / 10000) / YScaleQuantum)
                End If
                region.drawCursor Button, Shift, x, y
                
            Else
                'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & MinusInfinitySingle
                region.drawCursor Button, Shift, x, MinusInfinitySingle
            End If
        End If
    Next
    displayXAxisLabel x, 100
End If

mRegions(2 * index - 1).region.MouseMove Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim region As ChartRegion

Dim failpoint As Long
On Error GoTo Err

Set region = mRegions(2 * index - 1).region

If (region.snapCursorToTickBoundaries And Not CBool((Shift And vbCtrlMask))) Or _
    (Not region.snapCursorToTickBoundaries And CBool((Shift And vbCtrlMask))) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = region.YScaleQuantum
    If YScaleQuantum <> 0 Then y = YScaleQuantum * Int(y / YScaleQuantum)
End If

If Button = vbLeftButton Then mLeftDragging = False

region.MouseUp Button, Shift, x, y

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

lastVisiblePeriod = Round((CLng(HScroll.value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.currentPeriodNumber + chartWidth - 1))

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
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(x)
mLeftDragStartPosnY = y
mUserResizingRegions = True

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            Button As Integer, _
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
If Not mLeftDragging Then Exit Sub
If y = mLeftDragStartPosnY Then Exit Sub

' we resize the next region below the divider that has not
' been removed
For i = 2 * index + 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
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
    mRegions(currRegion).region.percentheight = 100 * newHeight / calcAvailableHeight
    mRegions(currRegion).percentheight = mRegions(currRegion).region.percentheight
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
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If Button = vbLeftButton Then mLeftDragging = False
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

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim failpoint As Long
On Error GoTo Err

Select Case Button.key
Case ToolbarCommandAutoScroll
    mAutoscroll = Not mAutoscroll
Case ToolbarCommandShowCrosshair
    pointerStyle = PointerCrosshairs
Case ToolbarCommandShowDiscCursor
    pointerStyle = PointerDisc
Case ToolbarCommandReduceSpacing
    If twipsPerBar >= 50 Then
        twipsPerBar = twipsPerBar - 25
    End If
    If twipsPerBar < 50 Then
        Button.Enabled = False
    End If
Case ToolbarCommandIncreaseSpacing
    twipsPerBar = twipsPerBar + 25
    Toolbar1.Buttons("reducespacing").Enabled = True
Case ToolbarCommandScrollLeft
    scrollX -(chartWidth * 0.2)
Case ToolbarCommandScrollRight
    scrollX chartWidth * 0.2
Case ToolbarCommandScrollEnd
    lastVisiblePeriod = currentPeriodNumber
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

mRegions(0).region.Click

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

mRegions(0).region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).region.MouseDown Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).region.MouseMove Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(0).region.MouseUp Button, Shift, x, y

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

mRegions(2 * index).region.Click

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

mRegions(2 * index).region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_DblClick" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).region.MouseDown Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseDown" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).region.MouseMove Button, Shift, x, y

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseMove" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
gLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).region.MouseUp Button, Shift, x, y

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
Dim region As ChartRegion
Dim ev As CollectionChangeEvent

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.addPeriod period.periodNumber, period.timestamp
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.addPeriod period.periodNumber, period.timestamp
If mSuppressDrawingCount = 0 Then setHorizontalScrollBar
setSession period.timestamp
If mAutoscroll Then scrollX 1

Set ev.affectedItem = period
ev.changeType = CollItemAdded
Set ev.Source = mPeriods
RaiseEvent PeriodsChanged(ev)
mController.firePeriodsChanged ev
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get allowHorizontalMouseScrolling() As Boolean
Attribute allowHorizontalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
allowHorizontalMouseScrolling = mAllowHorizontalMouseScrolling
End Property

Public Property Let allowHorizontalMouseScrolling(ByVal value As Boolean)
mAllowHorizontalMouseScrolling = value
PropertyChanged "allowHorizontalMouseScrolling"
End Property

Public Property Get allowVerticalMouseScrolling() As Boolean
Attribute allowVerticalMouseScrolling.VB_ProcData.VB_Invoke_Property = ";Behavior"
allowVerticalMouseScrolling = mAllowVerticalMouseScrolling
End Property

Public Property Let allowVerticalMouseScrolling(ByVal value As Boolean)
mAllowVerticalMouseScrolling = value
PropertyChanged "allowVerticalMouseScrolling"
End Property

Public Property Get autoscroll() As Boolean
Attribute autoscroll.VB_ProcData.VB_Invoke_Property = ";Behavior"
autoscroll = mAutoscroll
End Property

Public Property Let autoscroll(ByVal value As Boolean)
mAutoscroll = value
End Property

Public Property Let barTimePeriod( _
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

Public Property Get barTimePeriod() As TimePeriod
Set barTimePeriod = mBarTimePeriod
End Property

Public Property Get chartBackColor() As OLE_COLOR
Attribute chartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
chartBackColor = mDefaultRegionStyle.backColor
End Property

Public Property Let chartBackColor(ByVal val As OLE_COLOR)
Dim i As Long

If mDefaultRegionStyle.backColor = val Then Exit Property

UserControl.backColor = val

mDefaultRegionStyle.backColor = val
mXAxisRegion.regionBackColor = val

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        mRegions(i).region.regionBackColor = val
    End If
Next
paintAll
End Property

Public Property Get controller() As ChartController
Set controller = mController
End Property

Public Property Get chartLeft() As Single
chartLeft = mScaleLeft
End Property

Public Property Get chartWidth() As Single
chartWidth = YAxisPosition - mScaleLeft
End Property

Public Property Get currentPeriodNumber() As Long
currentPeriodNumber = mPeriods.currentPeriodNumber
End Property

Public Property Get currentSessionEndTime() As Date
currentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get currentSessionStartTime() As Date
currentSessionStartTime = mCurrentSessionStartTime
End Property

Public Property Get currentTool() As ToolTypes
currentTool = mCurrentTool
End Property

Public Property Let currentTool(ByVal value As ToolTypes)
Select Case value
Case ToolPointer
    mCurrentTool = value
Case ToolLine
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineExtended
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineRay
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineHorizontal
    mCurrentTool = ToolTypes.ToolPointer
Case ToolLineVertical
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciRetracement
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciExtension
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciCircle
    mCurrentTool = ToolTypes.ToolPointer
Case ToolFibonacciTime
    mCurrentTool = ToolTypes.ToolPointer
Case ToolRegressionChannel
    mCurrentTool = ToolTypes.ToolPointer
Case ToolRegressionEnvelope
    mCurrentTool = ToolTypes.ToolPointer
Case ToolText
    mCurrentTool = ToolTypes.ToolPointer
Case ToolPitchfork
    mCurrentTool = ToolTypes.ToolPointer
End Select
End Property

Public Property Get defaultBarDisplayMode() As BarDisplayModes
defaultBarDisplayMode = mDefaultBarDisplayMode
End Property

Public Property Let defaultBarDisplayMode( _
                ByVal value As BarDisplayModes)
mDefaultBarDisplayMode = value
End Property

Public Property Get defaultBarStyle() As BarStyle
Set defaultBarStyle = mDefaultBarStyle.clone
End Property

Public Property Let defaultBarStyle( _
                ByVal value As BarStyle)
Set mDefaultRegionStyle = value.clone
End Property

Public Property Get defaultDataPointStyle() As DataPointStyle
Set defaultDataPointStyle = mDefaultDataPointStyle.clone
End Property

Public Property Let defaultDataPointStyle( _
                ByVal value As DataPointStyle)
Set mDefaultDataPointStyle = value.clone
End Property

Public Property Get defaultLineStyle() As linestyle
Set defaultLineStyle = mDefaultLineStyle.clone
End Property

Public Property Let defaultLineStyle( _
                ByVal value As linestyle)
Set mDefaultLineStyle = value.clone
End Property

Public Property Get defaultRegionStyle() As ChartRegionStyle
Set defaultRegionStyle = mDefaultRegionStyle.clone
End Property

Public Property Let defaultRegionStyle( _
                ByVal value As ChartRegionStyle)
Set mDefaultRegionStyle = value.clone
End Property

Public Property Get defaultTextStyle() As TextStyle
Set defaultTextStyle = mDefaultTextStyle.clone
End Property

Public Property Let defaultTextStyle(ByVal value As TextStyle)
Set mDefaultTextStyle = value.clone
End Property

Public Property Get defaultYAxisStyle() As ChartRegionStyle
Set defaultYAxisStyle = mDefaultYAxisStyle.clone
End Property

Public Property Let defaultYAxisStyle(ByVal value As ChartRegionStyle)
Set mDefaultYAxisStyle = value.clone
End Property

Public Property Get firstVisiblePeriod() As Long
firstVisiblePeriod = mScaleLeft
End Property

Public Property Let firstVisiblePeriod(ByVal value As Long)
scrollX value - mScaleLeft + 1
End Property

Public Property Get lastVisiblePeriod() As Long
lastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let lastVisiblePeriod(ByVal value As Long)
scrollX value - mYAxisPosition + 1
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
Dim region As ChartRegion
mPointerCrosshairsColor = value
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.PointerCrosshairsColor = value
    End If
Next
End Property

Public Property Get PointerDiscColor() As OLE_COLOR
Attribute PointerDiscColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
PointerDiscColor = mPointerDiscColor
End Property

Public Property Let PointerDiscColor(ByVal value As OLE_COLOR)
Dim i As Long
Dim region As ChartRegion
mPointerDiscColor = value
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.PointerDiscColor = value
    End If
Next
End Property

Public Property Get pointerStyle() As PointerStyles
Attribute pointerStyle.VB_ProcData.VB_Invoke_Property = ";Appearance"
pointerStyle = mPointerStyle
End Property

Public Property Let pointerStyle(ByVal value As PointerStyles)
Dim i As Long
Dim region As ChartRegion
mPointerStyle = value
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.pointerStyle = value
    End If
Next
End Property

Public Property Get RegionDefaultAutoscale() As Boolean
Attribute RegionDefaultAutoscale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultAutoscale = mDefaultRegionStyle.autoscale
End Property

Public Property Let RegionDefaultAutoscale(ByVal value As Boolean)
mDefaultRegionStyle.autoscale = value
End Property

Public Property Get RegionDefaultBackColor() As OLE_COLOR
Attribute RegionDefaultBackColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultBackColor = mDefaultRegionStyle.backColor
End Property

Public Property Let RegionDefaultBackColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.backColor = val
End Property

Public Property Get RegionDefaultGridColor() As OLE_COLOR
Attribute RegionDefaultGridColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridColor = mDefaultRegionStyle.gridColor
End Property

Public Property Let RegionDefaultGridColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.gridColor = val
End Property

Public Property Get RegionDefaultGridlineSpacingY() As Double
Attribute RegionDefaultGridlineSpacingY.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridlineSpacingY = mDefaultRegionStyle.gridlineSpacingY
End Property

Public Property Let RegionDefaultGridlineSpacingY(ByVal value As Double)
mDefaultRegionStyle.gridlineSpacingY = value
End Property

Public Property Get RegionDefaultGridTextColor() As OLE_COLOR
Attribute RegionDefaultGridTextColor.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultGridTextColor = mDefaultRegionStyle.gridTextColor
End Property

Public Property Let RegionDefaultGridTextColor(ByVal val As OLE_COLOR)
mDefaultRegionStyle.gridTextColor = val
End Property

Public Property Get RegionDefaultHasGrid() As Boolean
Attribute RegionDefaultHasGrid.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGrid = mDefaultRegionStyle.hasGrid
End Property

Public Property Let RegionDefaultHasGrid(ByVal val As Boolean)
mDefaultRegionStyle.hasGrid = val
End Property

Public Property Get RegionDefaultHasGridText() As Boolean
Attribute RegionDefaultHasGridText.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultHasGridText = mDefaultRegionStyle.hasGridText
End Property

Public Property Let RegionDefaultHasGridText(ByVal val As Boolean)
mDefaultRegionStyle.hasGridText = val
End Property

Public Property Get RegionDefaultIntegerYScale() As Boolean
Attribute RegionDefaultIntegerYScale.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultIntegerYScale = mDefaultRegionStyle.integerYScale
End Property

Public Property Let RegionDefaultIntegerYScale(ByVal value As Boolean)
mDefaultRegionStyle.integerYScale = value
End Property

Public Property Get RegionDefaultMinimumHeight() As Double
Attribute RegionDefaultMinimumHeight.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultMinimumHeight = mDefaultRegionStyle.minimumHeight
End Property

Public Property Let RegionDefaultMinimumHeight(ByVal value As Double)
mDefaultRegionStyle.minimumHeight = value
End Property

Public Property Get RegionDefaultPointerStyle() As PointerStyles
Attribute RegionDefaultPointerStyle.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultPointerStyle = mDefaultRegionStyle.pointerStyle
End Property

Public Property Let RegionDefaultPointerStyle(ByVal value As PointerStyles)
mDefaultRegionStyle.pointerStyle = value
End Property

Public Property Get RegionDefaultYScaleQuantum() As Double
Attribute RegionDefaultYScaleQuantum.VB_ProcData.VB_Invoke_Property = ";Region Defaults"
RegionDefaultYScaleQuantum = mDefaultRegionStyle.YScaleQuantum
End Property

Public Property Let RegionDefaultYScaleQuantum(ByVal value As Double)
mDefaultRegionStyle.YScaleQuantum = value
End Property

Public Property Get sessionEndTime() As Date
Attribute sessionEndTime.VB_ProcData.VB_Invoke_Property = ";Behavior"
sessionEndTime = mSessionEndTime
End Property

Public Property Let sessionEndTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionEndTime", _
                "Value must be a time only"
mSessionEndTime = val
End Property

Public Property Get sessionStartTime() As Date
Attribute sessionStartTime.VB_ProcData.VB_Invoke_Property = ";Behavior"
sessionStartTime = mSessionStartTime
End Property

Public Property Let sessionStartTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionStartTime", _
                "Value must be a time only"
mSessionStartTime = val
End Property

Public Property Get showHorizontalScrollBar() As Boolean
Attribute showHorizontalScrollBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
showHorizontalScrollBar = mShowHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
mShowHorizontalScrollBar = val
If mShowHorizontalScrollBar Then
    'HScroll.height = HorizScrollBarHeight
    HScroll.visible = True
Else
    'HScroll.height = 0
    HScroll.visible = False
End If
Resize False, True
End Property

Public Property Get showToolbar() As Boolean
Attribute showToolbar.VB_ProcData.VB_Invoke_Property = ";Appearance"
showToolbar = mShowToolbar
End Property

Public Property Let showToolbar(ByVal val As Boolean)
mShowToolbar = val
If mShowToolbar Then
    Toolbar1.visible = True
Else
    Toolbar1.visible = False
End If
Resize False, True
End Property

Public Property Get suppressDrawing() As Boolean
suppressDrawing = (mSuppressDrawingCount > 0)
End Property

Public Property Let suppressDrawing(ByVal val As Boolean)
Dim i As Long
Dim region As ChartRegion
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
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.suppressDrawing = (mSuppressDrawingCount > 0)
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.suppressDrawing = (mSuppressDrawingCount > 0)
End Property

Public Property Get twipsPerBar() As Long
Attribute twipsPerBar.VB_ProcData.VB_Invoke_Property = ";Appearance"
twipsPerBar = mTwipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
mTwipsPerBar = val
resizeX
setHorizontalScrollBar
'paintAll
End Property

Public Property Set verticalGridTimePeriod( _
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

Public Property Get verticalGridTimePeriod() As TimePeriod
Set verticalGridTimePeriod = mVerticalGridTimePeriod
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

Public Function addChartRegion(ByVal percentheight As Double, _
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
Dim btn As Button

If name <> "" Then
    If Not getChartRegion(name) Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                "ChartSkil26.Chart::addChartRegion", _
                "Region " & name & " already exists"
    End If
End If

If style Is Nothing Then Set style = mDefaultRegionStyle
If yAxisStyle Is Nothing Then Set yAxisStyle = mDefaultYAxisStyle

Set addChartRegion = New ChartRegion
addChartRegion.name = name

If mRegionsIndex = 0 Then
    addChartRegion.toolbar = Toolbar1
    For Each btn In Toolbar1.Buttons
        btn.Enabled = True
        Select Case mPointerStyle
        Case PointerNone
            If btn.key = "showcrosshair" Then btn.value = tbrPressed
        Case PointerCrosshairs
            If btn.key = "showcrosshair" Then btn.value = tbrPressed
        Case PointerDisc
            If btn.key = "showdisccursor" Then btn.value = tbrPressed
        End Select
    Next
End If

mRegionsIndex = mRegionsIndex + 1
controlIndex = 1 + (mRegionsIndex - 1) / 2

Load ChartRegionPicture(controlIndex)
ChartRegionPicture(controlIndex).align = vbAlignNone
ChartRegionPicture(controlIndex).width = _
    UserControl.ScaleWidth * (mYAxisPosition - chartLeft) / XAxisPicture.ScaleWidth
ChartRegionPicture(controlIndex).visible = True

Load YAxisPicture(controlIndex)
YAxisPicture(controlIndex).align = vbAlignNone
YAxisPicture(controlIndex).left = ChartRegionPicture(controlIndex).width
YAxisPicture(controlIndex).width = UserControl.ScaleWidth - YAxisPicture(YAxisPicture.UBound).left
YAxisPicture(controlIndex).visible = True

addChartRegion.controller = controller
addChartRegion.canvas = createCanvas(ChartRegionPicture(controlIndex))
addChartRegion.suppressDrawing = (mSuppressDrawingCount > 0)
addChartRegion.currentTool = mCurrentTool
addChartRegion.minimumPercentHeight = minimumPercentHeight
addChartRegion.percentheight = percentheight
addChartRegion.pointerStyle = mPointerStyle
addChartRegion.PointerCrosshairsColor = mPointerCrosshairsColor
addChartRegion.PointerDiscColor = mPointerDiscColor
addChartRegion.regionLeft = mScaleLeft
addChartRegion.regionNumber = mRegionsIndex
addChartRegion.regionBottom = 0
addChartRegion.regionTop = 1
addChartRegion.periodsInView mScaleLeft, mYAxisPosition - 1
addChartRegion.verticalGridTimePeriod = mVerticalGridTimePeriod
addChartRegion.sessionStartTime = mSessionStartTime

addChartRegion.defaultBarStyle = mDefaultBarStyle
addChartRegion.defaultDataPointStyle = mDefaultDataPointStyle
addChartRegion.defaultLineStyle = mDefaultLineStyle
addChartRegion.defaultTextStyle = mDefaultTextStyle
addChartRegion.style = style

If mHideGrid Then addChartRegion.hideGrid


Set mRegions(mRegionsIndex).region = addChartRegion
If percentheight <> 100 Then
    mRegions(mRegionsIndex).percentheight = mRegionHeightReductionFactor * percentheight
Else
    mRegions(mRegionsIndex).useAvailableSpace = True
End If

Load RegionDividerPicture(controlIndex)
RegionDividerPicture(controlIndex).visible = True

mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then
    ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
End If

Set YAxisRegion = New ChartRegion
YAxisRegion.canvas = createCanvas(YAxisPicture(controlIndex))
YAxisRegion.regionNumber = mRegionsIndex
YAxisRegion.regionBottom = 0
YAxisRegion.regionTop = 1
YAxisRegion.isYAxisRegion = True
YAxisRegion.defaultBarStyle = mDefaultBarStyle
YAxisRegion.defaultDataPointStyle = mDefaultDataPointStyle
YAxisRegion.defaultLineStyle = mDefaultLineStyle
YAxisRegion.defaultTextStyle = mDefaultTextStyle
YAxisRegion.style = yAxisStyle
addChartRegion.YAxisRegion = YAxisRegion

Set mRegions(mRegionsIndex).region = YAxisRegion

mNumRegionsInUse = mNumRegionsInUse + 1

If sizeRegions Then
    Set ev.affectedItem = addChartRegion
    ev.changeType = CollItemAdded
    RaiseEvent RegionsChanged(ev)
    mController.fireRegionsChanged ev
    
'    ' now add all the current periods to ensure the grid lines are properly set up
'    ' NB: this might be a candidate for converting to a task for large charts
'    For Each var In mPeriods
'        Set p = var
'        addChartRegion.addPeriod p.periodNumber, p.timestamp
'    Next
Else
    ' can't fit this all in! So remove the added region,
    Set addChartRegion = Nothing
    Set mRegions(mRegionsIndex).region = Nothing
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

Public Function clearChart()
Dim i As Long
Dim controlIndex As Long

For i = 1 To mRegionsIndex Step 2
    controlIndex = 1 + (i - 1) / 2
    If Not mRegions(i).region Is Nothing Then
        mRegions(i).region.clearRegion
        mRegions(i + 1).region.clearRegion
        ChartRegionPicture(controlIndex).visible = False
        YAxisPicture(controlIndex).visible = False
        If i <> mRegionsIndex Then _
                RegionDividerPicture(controlIndex).visible = False
    End If
Next

Erase mRegions

If Not mXAxisRegion Is Nothing Then mXAxisRegion.clearRegion
XAxisPicture.Cls
Set mXAxisRegion = Nothing
mPeriods.finish
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

Public Sub displayGrid()
Dim i As Long
Dim region As ChartRegion

If Not mHideGrid Then Exit Sub

mHideGrid = False
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.displayGrid
    End If
Next
End Sub

Public Function getChartRegion(ByVal name As String) As ChartRegion
Dim i As Long

name = UCase$(name)
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        If UCase$(mRegions(i).region.name) = name Then
            Set getChartRegion = mRegions(i).region
            Exit Function
        End If
    End If
Next
                    
End Function

Public Sub hideGrid()
Dim i As Long
Dim region As ChartRegion

If mHideGrid Then Exit Sub

mHideGrid = True
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.hideGrid
    End If
Next
End Sub

Public Function isGridHidden() As Boolean
isGridHidden = mHideGrid
End Function

Public Function isTimeInSession(ByVal timestamp As Date) As Boolean

If timestamp >= mCurrentSessionStartTime And _
    timestamp < mCurrentSessionEndTime _
Then
    isTimeInSession = True
End If
End Function

Public Function refresh()
UserControl.refresh
End Function

Public Sub removeChartRegion( _
                    ByVal region As ChartRegion)
Dim i As Long
Dim ev As CollectionChangeEvent

If region.isXAxisRegion Or region.isYAxisRegion Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & "removeChartRegion", _
            "Cannot remove an axis region"
End If

For i = 1 To mRegionsIndex Step 2
    If region Is mRegions(i).region Then
        region.clearRegion
        Set mRegions(i).region = Nothing
        mRegions(i + 1).region.clearRegion
        Set mRegions(i + 1).region = Nothing
        RegionDividerPicture(1 + (i - 1) / 2).visible = False
        Exit For
    End If
Next

mNumRegionsInUse = mNumRegionsInUse - 1

sizeRegions
'paintAll

ev.changeType = CollItemRemoved
Set ev.affectedItem = region
mController.fireRegionsChanged ev
End Sub

Public Sub scrollX(ByVal value As Long)
Dim region As ChartRegion
Dim i As Long
If value = 0 Then Exit Sub

If (lastVisiblePeriod + value) > _
        (mPeriods.currentPeriodNumber + chartWidth - 1) Then
    value = mPeriods.currentPeriodNumber + chartWidth - 1 - lastVisiblePeriod
ElseIf (lastVisiblePeriod + value) < 1 Then
    value = 1 - lastVisiblePeriod
End If

mYAxisPosition = mYAxisPosition + value
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.width * mScaleWidth) - _
            mScaleWidth
XAxisPicture.ScaleLeft = mScaleLeft

If mSuppressDrawingCount > 0 Then Exit Sub

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        If Not mRegions(i).region Is Nothing Then
            Set region = mRegions(i).region
            region.periodsInView mScaleLeft, mYAxisPosition - 1
        End If
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.periodsInView mScaleLeft, mScaleLeft + mScaleWidth
setHorizontalScrollBar
'paintAll
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
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date)
Dim i As Long

i = -1
Do
    i = i + 1
Loop Until calcSessionTimesHelper(timestamp + i, sessionStartTime, sessionEndTime)
End Sub

Friend Function calcSessionTimesHelper(ByVal timestamp As Date, _
                            ByRef sessionStartTime As Date, _
                            ByRef sessionEndTime As Date) As Boolean
Dim referenceDate As Date
Dim referenceTime As Date
Dim weekday As Long

referenceDate = DateValue(timestamp)
referenceTime = TimeValue(timestamp)

If mSessionStartTime < mSessionEndTime Then
    ' session doesn't span midnight
    If referenceTime < mSessionEndTime Then
        sessionStartTime = referenceDate + mSessionStartTime
        sessionEndTime = referenceDate + mSessionEndTime
    Else
        sessionStartTime = referenceDate + 1 + mSessionStartTime
        sessionEndTime = referenceDate + 1 + mSessionEndTime
    End If
ElseIf mSessionStartTime > mSessionEndTime Then
    ' session spans midnight
    If referenceTime >= mSessionEndTime Then
        sessionStartTime = referenceDate + mSessionStartTime
        sessionEndTime = referenceDate + 1 + mSessionEndTime
    Else
        sessionStartTime = referenceDate - 1 + mSessionStartTime
        sessionEndTime = referenceDate + mSessionEndTime
    End If
Else
    ' this instrument trades 24hrs, or the contract service provider doesn't know
    ' the session start and end times
    sessionStartTime = referenceDate
    sessionEndTime = referenceDate + 1
End If

weekday = DatePart("w", sessionStartTime)
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
                ByVal surface As PictureBox) As canvas
Set createCanvas = New canvas
createCanvas.surface = surface
End Function

Private Sub createXAxisRegion()
Dim aFont As StdFont
Dim style As ChartRegionStyle

Set mXAxisRegion = New ChartRegion
mXAxisRegion.isXAxisRegion = True

Set mRegions(0).region = mXAxisRegion

mXAxisRegion.controller = controller
mXAxisRegion.canvas = createCanvas(XAxisPicture)
mXAxisRegion.verticalGridTimePeriod = mVerticalGridTimePeriod
mXAxisRegion.regionBottom = 0
mXAxisRegion.regionTop = 1
mXAxisRegion.sessionStartTime = mSessionStartTime

mXAxisRegion.defaultBarStyle = mDefaultBarStyle
mXAxisRegion.defaultDataPointStyle = mDefaultDataPointStyle
mXAxisRegion.defaultLineStyle = mDefaultLineStyle
mXAxisRegion.defaultTextStyle = mDefaultTextStyle

Set style = mDefaultRegionStyle.clone
style.hasGrid = False
style.hasGridText = True
style.pointerStyle = PointerNone
mXAxisRegion.style = style

Set mXCursorText = mXAxisRegion.addText(LayerNumbers.LayerPointer)
mXCursorText.align = AlignTopCentre
mXCursorText.Color = vbWhite Xor mDefaultRegionStyle.backColor
mXCursorText.box = True
mXCursorText.boxFillColor = vbWhite
mXCursorText.boxStyle = LineSolid
mXCursorText.boxColor = vbBlack
Set aFont = New StdFont
aFont.name = "Arial"
aFont.Size = 8
aFont.Underline = False
aFont.Bold = False
mXCursorText.font = aFont

End Sub

Private Sub displayXAxisLabel(x As Single, y As Single)
Dim thisPeriod As period
Dim periodNumber As Long
Dim prevPeriodNumber As Long
Dim prevPeriod As period

If mXAxisRegion Is Nothing Then createXAxisRegion

If Round(x) >= mYAxisPosition Then Exit Sub
If mPeriods.count = 0 Then Exit Sub

On Error Resume Next
periodNumber = Round(x)
Set thisPeriod = mPeriods(periodNumber)
On Error GoTo 0
If thisPeriod Is Nothing Then
    mXCursorText.text = ""
    Exit Sub
End If

mXCursorText.position = mXAxisRegion.newPoint( _
                            periodNumber, _
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
Dim btn As Button

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
    HScroll.visible = mShowHorizontalScrollBar
    Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
    mVerticalGridTimePeriodSet = False
    
    Set mDefaultRegionStyle = gCreateChartRegionStyle
    
    Set mDefaultBarStyle = gCreateBarStyle
    
    Set mDefaultDataPointStyle = gCreateDataPointStyle
    
    Set mDefaultLineStyle = gCreateLineStyle
    
    Set mDefaultTextStyle = gCreateTextStyle
    
    Set mDefaultYAxisStyle = gCreateChartRegionStyle
    mDefaultYAxisStyle.hasGrid = False
    
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

Private Sub paintAll()
Dim region As ChartRegion
Dim i As Long

If mSuppressDrawingCount > 0 Then Exit Sub

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.paintRegion
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.paintRegion

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
Dim region As ChartRegion

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
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.periodsInView mScaleLeft, mYAxisPosition - 1
    End If
Next

failpoint = 700

If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.periodsInView mScaleLeft, mScaleLeft + mScaleWidth
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

If mPeriods.currentPeriodNumber + chartWidth - 1 > 32767 Then

    failpoint = 100

    HScroll.Max = 32767
ElseIf mPeriods.currentPeriodNumber + chartWidth - 1 < 1 Then

    failpoint = 200

    HScroll.Max = 1
Else

    failpoint = 300
    
    HScroll.Max = mPeriods.currentPeriodNumber + chartWidth - 1
End If
HScroll.Min = 0


failpoint = 400

' NB the following calculation has to be done using doubles as for very large charts it can cause an overflow using integers
hscrollVal = Round(CDbl(HScroll.Max) * CDbl(lastVisiblePeriod) / CDbl((mPeriods.currentPeriodNumber + chartWidth - 1)))
If hscrollVal > HScroll.Max Then
    HScroll.value = HScroll.Max
ElseIf hscrollVal < HScroll.Min Then
    HScroll.value = HScroll.Min
Else
    HScroll.value = Round(CDbl(HScroll.Max) * CDbl(lastVisiblePeriod) / CDbl((mPeriods.currentPeriodNumber + chartWidth - 1)))
End If

failpoint = 500

HScroll.SmallChange = 1
If (chartWidth - 1) < 1 Then
    HScroll.LargeChange = 1
Else
    HScroll.LargeChange = chartWidth - 1
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
Dim region As ChartRegion
mXAxisRegion.verticalGridTimePeriod = mVerticalGridTimePeriod
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).region Is Nothing Then
        Set region = mRegions(i).region
        region.verticalGridTimePeriod = mVerticalGridTimePeriod
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
    If Not mRegions(i).region Is Nothing Then
        Set aRegion = mRegions(i).region
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
        If Not mRegions(i).region Is Nothing Then
            Set aRegion = mRegions(i).region
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
    If Not mRegions(i).region Is Nothing Then
        Set aRegion = mRegions(i).region
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
    If Not mRegions(i).region Is Nothing Then
        Set aRegion = mRegions(i).region
        If Not suppressDrawing Then
            ChartRegionPicture(controlIndex).height = mRegions(i).actualHeight
            YAxisPicture(controlIndex).height = mRegions(i).actualHeight
            ChartRegionPicture(controlIndex).top = top
            YAxisPicture(controlIndex).top = top
            aRegion.resizedY
        End If
        top = top + mRegions(i).actualHeight
        numRegionsSized = numRegionsSized + 1
        If Not suppressDrawing Then
            RegionDividerPicture(controlIndex).top = top
        End If
        If numRegionsSized <> mNumRegionsInUse Then
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbSizeNS
        Else
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbDefault
        End If
        top = top + RegionDividerPicture(controlIndex).height
    Else
        If Not suppressDrawing Then
            ChartRegionPicture(controlIndex).visible = False
            YAxisPicture(controlIndex).visible = False
            RegionDividerPicture(controlIndex).visible = False
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

