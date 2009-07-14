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
   Begin VB.PictureBox SelectorPicture 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2640
      Picture         =   "ChartArea.ctx":0000
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox BlankPicture 
      AutoSize        =   -1  'True
      Height          =   540
      Left            =   2280
      Picture         =   "ChartArea.ctx":0152
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
            Picture         =   "ChartArea.ctx":0594
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":09E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":0E38
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":128A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":16DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1B2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":1F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":23D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2824
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":2C76
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":30C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":351A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":396C
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":3DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4210
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4662
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":4AB4
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
            Picture         =   "ChartArea.ctx":4F06
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5358
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":57AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":5BFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":604E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":64A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":68F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":6D44
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7196
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":75E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7A3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":7E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":82DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8730
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":8FD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9426
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
      BackColor       =   &H0000C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   0
      MouseIcon       =   "ChartArea.ctx":9878
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
            Picture         =   "ChartArea.ctx":9CBA
            Key             =   "showbars"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9FD4
            Key             =   "showcandlesticks"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A2EE
            Key             =   "showline"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A608
            Key             =   "showcrosshair"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":A922
            Key             =   "showdisccursor"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":AC3C
            Key             =   "thinnerbars"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":AF56
            Key             =   "thickerbars"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B270
            Key             =   "narrower"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B6C2
            Key             =   "wider"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":B9DC
            Key             =   "scaledown"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":BCF6
            Key             =   "scaleup"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C010
            Key             =   "scrolldown"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C32A
            Key             =   "scrollup"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C644
            Key             =   "scrollleft"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":C95E
            Key             =   "scrollright"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CC78
            Key             =   "scrollend"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":CF92
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
            Picture         =   "ChartArea.ctx":D2AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D5C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":D8E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DBFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":DF14
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E22E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E548
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":E862
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":ECB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F106
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F420
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":F73A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":FA54
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":FD6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":10088
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":103A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":106BC
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
Event MouseDown(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseDown.VB_UserMemId = -605
                
Event mouseMove(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute mouseMove.VB_UserMemId = -606
                
Event MouseUp(Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Event PointerModeChanged()
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
    PercentHeight       As Double
    actualHeight        As Long
    useAvailableSpace   As Boolean
End Type

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                                As String = "Chart"

Private Const PropNameHorizontalMouseScrollingAllowed     As String = "HorizontalMouseScrollingAllowed"
Private Const PropNameVerticalMouseScrollingAllowed       As String = "VerticalMouseScrollingAllowed"
Private Const PropNameAutoscrolling                        As String = "Autoscrolling"
Private Const PropNameChartBackColor                    As String = "ChartBackColor"
Private Const PropNamePeriodLength                      As String = "PeriodLength"
Private Const PropNamePeriodUnits                       As String = "PeriodUnits"
Private Const PropNamePointerDiscColor                  As String = "PointerDiscColor"
Private Const PropNamePointerCrosshairsColor            As String = "PointerCrosshairsColor"
Private Const PropNamePointerStyle                      As String = "PointerStyle"
Private Const PropNameHorizontalScrollBarVisible           As String = "HorizontalScrollBarVisible"
Private Const PropNameToolbarVisible                       As String = "ShowToobar"
Private Const PropNameTwipsPerBar                       As String = "TwipsPerBar"
Private Const PropNameVerticalGridSpacing               As String = "VerticalGridSpacing"
Private Const PropNameVerticalGridUnits                 As String = "VerticalGridUnits"
Private Const PropNameXAxisVisible                      As String = "XAxisVisible"
Private Const PropNameYAxisVisible                      As String = "YAxisVisible"
Private Const PropNameYAxisWidthCm                      As String = "YAxisWidthCm"

Private Const PropDfltHorizontalMouseScrollingAllowed     As Boolean = True
Private Const PropDfltVerticalMouseScrollingAllowed       As Boolean = True
Private Const PropDfltAutoscrolling                        As Boolean = True
Private Const PropDfltChartBackColor                    As Long = &H643232
Private Const PropDfltPeriodLength                      As Long = 5
Private Const PropDfltPeriodUnits                       As Long = TimePeriodMinute
Private Const PropDfltPointerDiscColor                  As Long = &H89FFFF
Private Const PropDfltPointerCrosshairsColor            As Long = &HC1DFE
Private Const PropDfltPointerStyle                      As Long = PointerStyles.PointerCrosshairs
Private Const PropDfltHorizontalScrollBarVisible           As Boolean = True
Private Const PropDfltToolbarVisible                       As Boolean = True
Private Const PropDfltTwipsPerBar                       As Long = 150
Private Const PropDfltVerticalGridSpacing               As Long = 1
Private Const PropDfltVerticalGridUnits                 As Long = TimePeriodHour
Private Const PropDfltXAxisVisible                      As Boolean = True
Private Const PropDfltYAxisVisible                      As Boolean = True
Private Const PropDfltYAxisWidthCm                      As Single = 1.3

'================================================================================
' Member variables
'================================================================================

Private mRegions() As RegionTableEntry
Private mRegionsIndex As Long
Private mNumRegionsInUse As Long

Private WithEvents mPeriods As Periods
Attribute mPeriods.VB_VarHelpID = -1

Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single
Private mPrevWidth As Single

Private mTwipsPerBar As Long

Private mXAxisVisible  As Boolean
Private mXAxisRegion As ChartRegion
Private mXCursorText As Text

Private mYAxisPosition As Long
Private mYAxisWidthCm As Single
Private mYAxisVisible  As Boolean

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
Private mToolPointerStyle As PointerStyles
Private mToolIcon As IPictureDisp
Private mPointerCrosshairsColor As Long
Private mPointerDiscColor As Long

Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mSuppressDrawingCount As Long
Private mPainted As Boolean

Private mLeftDragStartPosnX As Long
Private mLeftDragStartPosnY As Single

Private mUserResizingRegions As Boolean

Private mHorizontalMouseScrollingAllowed As Boolean
Private mVerticalMouseScrollingAllowed As Boolean

Private mMouseScrollingInProgress As Boolean

Private mHorizontalScrollBarVisible As Boolean
Private mToolbarVisible As Boolean

Private mRegionHeightReductionFactor As Double

Private mReferenceTime As Date

Private mAutoscrolling As Boolean

Private mBackGroundCanvas As Canvas
Private mChartBackGradientFillColors() As Long

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
Dim failpoint As Long
On Error GoTo Err

Set gBlankCursor = BlankPicture.Picture
Set gSelectorCursor = SelectorPicture.Picture

ReDim mChartBackGradientFillColors(0) As Long
mChartBackGradientFillColors(0) = PropDfltChartBackColor

Set mBackGroundCanvas = New Canvas
mBackGroundCanvas.Surface = ChartRegionPicture(0)
mBackGroundCanvas.MousePointer = vbDefault

Initialise
createXAxisRegion

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "UserControl_Initialize" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_InitProperties()
On Error Resume Next

HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
Autoscrolling = PropDfltAutoscrolling
'mXAxisRegion.BackColor = PropDfltChartBackColor
PointerCrosshairsColor = PropDfltPointerCrosshairsColor
PointerDiscColor = PropDfltPointerDiscColor
PointerStyle = PropDfltPointerStyle
HorizontalScrollBarVisible = PropDfltHorizontalScrollBarVisible
ToolbarVisible = PropDfltToolbarVisible
TwipsPerBar = PropDfltTwipsPerBar
'Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
XAxisVisible = PropDfltXAxisVisible
YAxisWidthCm = PropDfltYAxisWidthCm
YAxisVisible = PropDfltYAxisVisible

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

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent mouseMove(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(X, vbTwips, vbContainerPosition), _
                    ScaleY(Y, vbTwips, vbContainerPosition))
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

On Error Resume Next

HorizontalMouseScrollingAllowed = PropBag.ReadProperty(PropNameHorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed)
If Err.Number <> 0 Then
    HorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
    Err.Clear
End If

VerticalMouseScrollingAllowed = PropBag.ReadProperty(PropNameVerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed)
If Err.Number <> 0 Then
    VerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed
    Err.Clear
End If

Autoscrolling = PropBag.ReadProperty(PropNameAutoscrolling, PropDfltAutoscrolling)
If Err.Number <> 0 Then
    Autoscrolling = PropDfltAutoscrolling
    Err.Clear
End If

ChartBackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    ChartBackColor = PropDfltChartBackColor
    Err.Clear
End If

mXAxisRegion.BackColor = PropBag.ReadProperty(PropNameChartBackColor, PropDfltChartBackColor)
If Err.Number <> 0 Then
    mXAxisRegion.BackColor = PropDfltChartBackColor
    Err.Clear
End If

PointerCrosshairsColor = PropBag.ReadProperty(PropNamePointerCrosshairsColor, PropDfltPointerCrosshairsColor)
If Err.Number <> 0 Then
    PointerCrosshairsColor = PropDfltPointerCrosshairsColor
    Err.Clear
End If

PointerDiscColor = PropBag.ReadProperty(PropNamePointerDiscColor, PropDfltPointerDiscColor)
If Err.Number <> 0 Then
    PointerDiscColor = PropDfltPointerDiscColor
    Err.Clear
End If

PointerStyle = PropBag.ReadProperty(PropNamePointerStyle, PropDfltPointerStyle)
If Err.Number <> 0 Then
    PointerStyle = PropDfltPointerStyle
    Err.Clear
End If

HorizontalScrollBarVisible = PropBag.ReadProperty(PropNameHorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible)
If Err.Number <> 0 Then
    HorizontalScrollBarVisible = PropDfltHorizontalScrollBarVisible
    Err.Clear
End If

ToolbarVisible = PropBag.ReadProperty(PropNameToolbarVisible, PropDfltToolbarVisible)
If Err.Number <> 0 Then
    ToolbarVisible = PropDfltToolbarVisible
    Err.Clear
End If

TwipsPerBar = PropBag.ReadProperty(PropNameTwipsPerBar, PropDfltTwipsPerBar)
If Err.Number <> 0 Then
    TwipsPerBar = PropDfltTwipsPerBar
    Err.Clear
End If

Set mVerticalGridTimePeriod = GetTimePeriod(PropBag.ReadProperty(PropNameVerticalGridSpacing, PropDfltVerticalGridSpacing), _
                        PropBag.ReadProperty(PropNameVerticalGridUnits, PropDfltVerticalGridUnits))
If Err.Number <> 0 Then
    Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
    Err.Clear
End If

XAxisVisible = PropBag.ReadProperty(PropNameXAxisVisible, PropDfltXAxisVisible)
If Err.Number <> 0 Then
    XAxisVisible = PropDfltXAxisVisible
    Err.Clear
End If

YAxisWidthCm = PropBag.ReadProperty(PropNameYAxisWidthCm, PropDfltYAxisWidthCm)
If Err.Number <> 0 Then
    YAxisWidthCm = PropDfltYAxisWidthCm
    Err.Clear
End If

YAxisVisible = PropBag.ReadProperty(PropNameYAxisVisible, PropDfltYAxisVisible)
If Err.Number <> 0 Then
    YAxisVisible = PropDfltYAxisVisible
    Err.Clear
End If

Initialise

End Sub

Private Sub UserControl_Resize()
Static resizeCount As Long

Dim failpoint As Long
On Error GoTo Err

'gLogger.Log LogLevelDetail, "ChartSkil: UserControl_Resize: enter"
resizeCount = resizeCount + 1

Resize (UserControl.Width <> mPrevWidth), (UserControl.Height <> mPrevHeight)
mPrevHeight = UserControl.Height
mPrevWidth = UserControl.Width

'gLogger.Log LogLevelDetail, "ChartSkil: UserControl_Resize: exit"

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "UserControl_Resize" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

Private Sub UserControl_Terminate()
Debug.Print "ChartSkil Usercontrol terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty PropNameHorizontalMouseScrollingAllowed, HorizontalMouseScrollingAllowed, PropDfltHorizontalMouseScrollingAllowed
PropBag.WriteProperty PropNameVerticalMouseScrollingAllowed, VerticalMouseScrollingAllowed, PropDfltVerticalMouseScrollingAllowed
PropBag.WriteProperty PropNameAutoscrolling, Autoscrolling, PropDfltAutoscrolling
If UBound(mChartBackGradientFillColors) = 0 Then
    PropBag.WriteProperty PropNameChartBackColor, mChartBackGradientFillColors(0)
End If
PropBag.WriteProperty PropNamePeriodLength, mBarTimePeriod.length, PropDfltPeriodLength
PropBag.WriteProperty PropNamePeriodUnits, mBarTimePeriod.units, PropDfltPeriodUnits
PropBag.WriteProperty PropNamePointerCrosshairsColor, PointerCrosshairsColor, PropDfltPointerCrosshairsColor
PropBag.WriteProperty PropNamePointerDiscColor, PointerDiscColor, PropDfltPointerDiscColor
PropBag.WriteProperty PropNamePointerStyle, mPointerStyle, PropDfltPointerStyle
PropBag.WriteProperty PropNameHorizontalScrollBarVisible, HorizontalScrollBarVisible, PropDfltHorizontalScrollBarVisible
PropBag.WriteProperty PropNameToolbarVisible, ToolbarVisible, PropDfltToolbarVisible
PropBag.WriteProperty PropNameTwipsPerBar, TwipsPerBar, PropDfltTwipsPerBar
PropBag.WriteProperty PropNameVerticalGridSpacing, mVerticalGridTimePeriod.length, PropDfltVerticalGridSpacing
PropBag.WriteProperty PropNameVerticalGridUnits, mVerticalGridTimePeriod.units, PropDfltVerticalGridUnits
PropBag.WriteProperty PropNameXAxisVisible, XAxisVisible, PropDfltXAxisVisible
PropBag.WriteProperty PropNameYAxisVisible, YAxisVisible, PropDfltYAxisVisible
PropBag.WriteProperty PropNameYAxisWidthCm, YAxisWidthCm, PropDfltYAxisWidthCm
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_Click(index As Integer)
Dim Region As ChartRegion
Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

Set Region = mRegions(2 * index - 1).Region
Region.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_Click" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

mRegions(2 * index - 1).Region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_DblClick" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Dim failpoint As Long

On Error GoTo Err

If index = 0 Then Exit Sub

Set Region = mRegions(2 * index - 1).Region


If CBool(Button And MouseButtonConstants.vbLeftButton) Then mMouseScrollingInProgress = True

' we notify the region selection first so that the application has a chance to
' turn off scrolling and snapping before getting the MouseDown event
RaiseEvent RegionSelected(Region)

If (mPointerMode = PointerModeDefault And _
        ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int((Y + YScaleQuantum / 10000) / YScaleQuantum)
End If

If mPointerMode = PointerModeDefault And _
    (mHorizontalMouseScrollingAllowed Or mVerticalMouseScrollingAllowed) _
Then
    mLeftDragStartPosnX = Int(X)
    mLeftDragStartPosnY = Y
End If

Region.MouseDown Button, Shift, Round(X), Y
RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(ChartRegionPicture(index).Left + ChartRegionPicture(index).ScaleX(X, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(ChartRegionPicture(index).Top + ChartRegionPicture(index).ScaleY(Y, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseDown" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                X As Single, _
                                Y As Single)

Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

If CBool(Button And MouseButtonConstants.vbLeftButton) Then
    If mPointerMode = PointerModeDefault And _
        (mHorizontalMouseScrollingAllowed Or mVerticalMouseScrollingAllowed) And _
        mMouseScrollingInProgress _
    Then
        mouseScroll index, Button, Shift, X, Y
    Else
        mMouseScrollingInProgress = False
        mouseMove index, Button, Shift, X, Y
    End If
Else
    mouseMove index, Button, Shift, X, Y
End If

mRegions(2 * index - 1).Region.mouseMove Button, Shift, Round(X), Y

RaiseEvent mouseMove(Button, _
                    Shift, _
                    ScaleX(ChartRegionPicture(index).Left + ChartRegionPicture(index).ScaleX(X, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(ChartRegionPicture(index).Top + ChartRegionPicture(index).ScaleY(Y, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseMove" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err

If index = 0 Then Exit Sub

mMouseScrollingInProgress = False

Set Region = mRegions(2 * index - 1).Region

If (mPointerMode = PointerModeDefault And _
        ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
        (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
    (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
Then
    Dim YScaleQuantum As Double
    YScaleQuantum = Region.YScaleQuantum
    If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int(Y / YScaleQuantum)
End If

Region.MouseUp Button, Shift, Round(X), Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(ChartRegionPicture(index).Left + ChartRegionPicture(index).ScaleX(X, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(ChartRegionPicture(index).Top + ChartRegionPicture(index).ScaleY(Y, ChartRegionPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ChartRegionPicture_MouseUp" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
Dim failpoint As Long
On Error GoTo Err

LastVisiblePeriod = Round((CLng(HScroll.value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.CurrentPeriodNumber + ChartWidth - 1))

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "HScroll_Change" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' RegionDividerPicture Event Handlers
'================================================================================

Private Sub RegionDividerPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If CBool(Button And MouseButtonConstants.vbLeftButton) Then
    mLeftDragStartPosnX = Int(X)
    mLeftDragStartPosnY = Y
    mUserResizingRegions = True
End If
RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(RegionDividerPicture(index).Left + RegionDividerPicture(index).ScaleX(X, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(RegionDividerPicture(index).Top + RegionDividerPicture(index).ScaleY(Y, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseDown" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseMove( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim vertChange As Long
Dim currRegion As Long
Dim newHeight As Long
Dim prevPercentHeight As Double
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
If Not CBool(Button And MouseButtonConstants.vbLeftButton) Then Exit Sub
If Y = mLeftDragStartPosnY Then Exit Sub

' we resize the next region below the divider that has not
' been removed
For i = 2 * index + 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        currRegion = i
        Exit For
    End If
Next

vertChange = mLeftDragStartPosnY - Y
newHeight = mRegions(currRegion).actualHeight + vertChange
If newHeight < 0 Then newHeight = 0

' the region table indicates the requested percentage used by each region
' and the actual Height allocation. We need to work out the new percentage
' for the region to be resized.

'prevPercentHeight = mRegions(currRegion).region.PercentHeight

prevPercentHeight = mRegions(currRegion).PercentHeight
If Not mRegions(currRegion).useAvailableSpace Then
    'mRegions(currRegion).region.PercentHeight = prevPercentHeight * newHeight / mRegions(currRegion).actualHeight
    mRegions(currRegion).PercentHeight = 100 * newHeight / calcAvailableHeight
    'mRegions(currRegion).PercentHeight = prevPercentHeight * newHeight / mRegions(currRegion).actualHeight
Else
    ' this is a 'use available space' region that's being resized. Now change
    ' it to use a specific percentage
    mRegions(currRegion).Region.PercentHeight = 100 * newHeight / calcAvailableHeight
    mRegions(currRegion).PercentHeight = mRegions(currRegion).Region.PercentHeight
End If

If sizeRegions Then
    'paintAll
Else
    ' the regions couldn't be resized so reset the region's percent Height
    mRegions(currRegion).PercentHeight = prevPercentHeight
End If

RaiseEvent mouseMove(Button, _
                    Shift, _
                    ScaleX(RegionDividerPicture(index).Left + RegionDividerPicture(index).ScaleX(X, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(RegionDividerPicture(index).Top + RegionDividerPicture(index).ScaleY(Y, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseMove" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            X As Single, _
                            Y As Single)
Dim failpoint As Long
On Error GoTo Err

If index = mRegionsIndex + 1 Then Exit Sub
mUserResizingRegions = False

RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(RegionDividerPicture(index).Left + RegionDividerPicture(index).ScaleX(X, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(RegionDividerPicture(index).Top + RegionDividerPicture(index).ScaleY(Y, RegionDividerPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RegionDividerPicture_MouseUp" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Toolbar1 Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim failpoint As Long
On Error GoTo Err

Select Case Button.Key
Case ToolbarCommandAutoscroll
    mAutoscrolling = Not mAutoscrolling
Case ToolbarCommandShowCrosshair
    PointerStyle = PointerCrosshairs
Case ToolbarCommandShowDiscCursor
    PointerStyle = PointerDisc
Case ToolbarCommandReduceSpacing
    If TwipsPerBar >= 50 Then
        TwipsPerBar = TwipsPerBar - 25
    End If
    If TwipsPerBar < 50 Then
        Button.Enabled = False
    End If
Case ToolbarCommandIncreaseSpacing
    TwipsPerBar = TwipsPerBar + 25
    Toolbar1.Buttons("reducespacing").Enabled = True
Case ToolbarCommandScrollLeft
    ScrollX -(ChartWidth * 0.2)
Case ToolbarCommandScrollRight
    ScrollX ChartWidth * 0.2
Case ToolbarCommandScrollEnd
    LastVisiblePeriod = CurrentPeriodNumber
End Select

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Toolbar1_ButtonClick" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource

End Sub

'================================================================================
' XAxisPicture Event Handlers
'================================================================================

Private Sub XAxisPicture_Click()
Dim failpoint As Long
On Error GoTo Err

mXAxisRegion.Click

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_Click" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_DblClick()
Dim failpoint As Long
On Error GoTo Err

mXAxisRegion.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_DblClick" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mXAxisRegion.MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(XAxisPicture.Left + XAxisPicture.ScaleX(X, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(XAxisPicture.Top + XAxisPicture.ScaleY(Y, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseDown" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mXAxisRegion.mouseMove Button, Shift, X, Y

RaiseEvent mouseMove(Button, _
                    Shift, _
                    ScaleX(XAxisPicture.Left + XAxisPicture.ScaleX(X, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(XAxisPicture.Top + XAxisPicture.ScaleY(Y, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseMove" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub XAxisPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mXAxisRegion.MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(XAxisPicture.Left + XAxisPicture.ScaleX(X, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(XAxisPicture.Top + XAxisPicture.ScaleY(Y, XAxisPicture.ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "XAxisPicture_MouseUp" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_Click" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_DblClick(index As Integer)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.DblCLick

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_DblClick" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.MouseDown Button, Shift, X, Y

RaiseEvent MouseDown(Button, _
                    Shift, _
                    ScaleX(YAxisPicture(index).Left + YAxisPicture(index).ScaleX(X, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(YAxisPicture(index).Top + YAxisPicture(index).ScaleY(Y, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseDown" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.mouseMove Button, Shift, X, Y

RaiseEvent mouseMove(Button, _
                    Shift, _
                    ScaleX(YAxisPicture(index).Left + YAxisPicture(index).ScaleX(X, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(YAxisPicture(index).Top + YAxisPicture(index).ScaleY(Y, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseMove" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

Private Sub YAxisPicture_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim failpoint As Long
On Error GoTo Err

mRegions(2 * index).Region.MouseUp Button, Shift, X, Y

RaiseEvent MouseUp(Button, _
                    Shift, _
                    ScaleX(YAxisPicture(index).Left + YAxisPicture(index).ScaleX(X, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition), _
                    ScaleY(YAxisPicture(index).Top + YAxisPicture(index).ScaleY(Y, YAxisPicture(index).ScaleMode, vbTwips), vbTwips, vbContainerPosition))
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "YAxisPicture_MouseUp" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' mPeriods Event Handlers
'================================================================================

Private Sub mPeriods_PeriodAdded(ByVal Period As Period)
Dim i As Long
Dim Region As ChartRegion
Dim ev As CollectionChangeEvent

Dim failpoint As Long
On Error GoTo Err

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.addPeriod Period.PeriodNumber, Period.Timestamp
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.addPeriod Period.PeriodNumber, Period.Timestamp
If mSuppressDrawingCount = 0 Then setHorizontalScrollBar
setSession Period.Timestamp
If mAutoscrolling Then ScrollX 1

Set ev.affectedItem = Period
ev.changeType = CollItemAdded
Set ev.Source = mPeriods
RaiseEvent PeriodsChanged(ev)

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mPeriods_PeriodAdded" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
End Sub

'================================================================================
' Properties
'================================================================================

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
Attribute BarTimePeriod.VB_MemberFlags = "400"
Set BarTimePeriod = mBarTimePeriod
End Property

Public Property Get Autoscrolling() As Boolean
Autoscrolling = mAutoscrolling
End Property

Public Property Let Autoscrolling(ByVal value As Boolean)
mAutoscrolling = value
PropertyChanged PropNameAutoscrolling
End Property

Public Property Get ChartBackColor() As OLE_COLOR
Attribute ChartBackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
ChartBackColor = mChartBackGradientFillColors(0)
End Property

Public Property Let ChartBackColor(ByVal value As OLE_COLOR)
ReDim mChartBackGradientFillColors(0) As Long
mChartBackGradientFillColors(0) = value
resizeBackground
PropertyChanged PropNameChartBackColor
End Property

Public Property Get ChartBackGradientFillColors() As Long()
Attribute ChartBackGradientFillColors.VB_ProcData.VB_Invoke_Property = ";Appearance"
ChartBackGradientFillColors = mChartBackGradientFillColors
End Property

Public Property Let ChartBackGradientFillColors(ByRef value() As Long)
Dim ar() As Long
ar = value
mChartBackGradientFillColors = ar
resizeBackground
PropertyChanged PropNameChartBackColor
End Property

Public Property Get ChartLeft() As Double
Attribute ChartLeft.VB_MemberFlags = "400"
ChartLeft = mScaleLeft
End Property

Public Property Get ChartWidth() As Double
Attribute ChartWidth.VB_MemberFlags = "400"
ChartWidth = YAxisPosition - mScaleLeft
End Property

Public Property Get CurrentPeriodNumber() As Long
Attribute CurrentPeriodNumber.VB_MemberFlags = "400"
CurrentPeriodNumber = mPeriods.CurrentPeriodNumber
End Property

Public Property Get CurrentSessionEndTime() As Date
Attribute CurrentSessionEndTime.VB_MemberFlags = "400"
CurrentSessionEndTime = mCurrentSessionEndTime
End Property

Public Property Get CurrentSessionStartTime() As Date
Attribute CurrentSessionStartTime.VB_MemberFlags = "400"
CurrentSessionStartTime = mCurrentSessionStartTime
End Property

Public Property Get FirstVisiblePeriod() As Long
Attribute FirstVisiblePeriod.VB_MemberFlags = "400"
FirstVisiblePeriod = mScaleLeft
End Property

Public Property Let FirstVisiblePeriod(ByVal value As Long)
ScrollX value - mScaleLeft + 1
End Property

Public Property Get HorizontalMouseScrollingAllowed() As Boolean
HorizontalMouseScrollingAllowed = mHorizontalMouseScrollingAllowed
End Property

Public Property Let HorizontalMouseScrollingAllowed(ByVal value As Boolean)
mHorizontalMouseScrollingAllowed = value
PropertyChanged PropNameHorizontalMouseScrollingAllowed
End Property

Public Property Get HorizontalScrollBarVisible() As Boolean
HorizontalScrollBarVisible = mHorizontalScrollBarVisible
PropertyChanged PropNameHorizontalScrollBarVisible
End Property

Public Property Let HorizontalScrollBarVisible(ByVal val As Boolean)
mHorizontalScrollBarVisible = val
If mHorizontalScrollBarVisible Then
    HScroll.Visible = True
Else
    HScroll.Visible = False
End If
Resize False, True
End Property

Public Property Get IsDrawingEnabled() As Boolean
IsDrawingEnabled = (mSuppressDrawingCount > 0)
End Property

Public Property Get IsGridHidden() As Boolean
IsGridHidden = mHideGrid
End Property

Public Property Get LastVisiblePeriod() As Long
Attribute LastVisiblePeriod.VB_MemberFlags = "400"
LastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let LastVisiblePeriod(ByVal value As Long)
ScrollX value - mYAxisPosition + 1
End Property

Public Property Get Periods() As Periods
Attribute Periods.VB_MemberFlags = "400"
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
PropertyChanged PropNamePointerCrosshairsColor
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
PropertyChanged PropNamePointerDiscColor
End Property

Public Property Get PointerIcon() As IPictureDisp
Attribute PointerIcon.VB_MemberFlags = "400"
Set PointerIcon = mPointerIcon
End Property

Public Property Let PointerIcon(ByVal value As IPictureDisp)
Dim i As Long
Dim Region As ChartRegion

If value Is Nothing Then Exit Property
If value Is mPointerIcon Then Exit Property

Set mPointerIcon = value

If mPointerStyle = PointerCustom Then
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            Set Region = mRegions(i).Region
            Region.PointerIcon = value
            Region.PointerStyle = PointerCustom
        End If
    Next
End If
End Property

Public Property Get PointerMode() As PointerModes
Attribute PointerMode.VB_MemberFlags = "400"
PointerMode = mPointerMode
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

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        If mPointerStyle = PointerCustom Then Region.PointerIcon = mPointerIcon
        Region.PointerStyle = mPointerStyle
    End If
Next
PropertyChanged PropNamePointerStyle
End Property

Public Property Get SessionEndTime() As Date
Attribute SessionEndTime.VB_MemberFlags = "400"
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
Attribute SessionStartTime.VB_MemberFlags = "400"
SessionStartTime = mSessionStartTime
End Property

Public Property Let SessionStartTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
                "ChartSkil26.Chart::(Let)sessionStartTime", _
                "Value must be a time only"
mSessionStartTime = val
End Property

Public Property Get ToolbarVisible() As Boolean
Attribute ToolbarVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
ToolbarVisible = mToolbarVisible
End Property

Public Property Let ToolbarVisible(ByVal val As Boolean)
mToolbarVisible = val
If mToolbarVisible Then
    Toolbar1.Visible = True
Else
    Toolbar1.Visible = False
End If
Resize False, True
PropertyChanged PropNameToolbarVisible
End Property

Public Property Let TwipsPerBar(ByVal val As Long)
Attribute TwipsPerBar.VB_ProcData.VB_Invoke_PropertyPut = ";Appearance"
mTwipsPerBar = val
resizeX
setHorizontalScrollBar
'paintAll
PropertyChanged PropNameTwipsPerBar
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
Attribute VerticalGridTimePeriod.VB_MemberFlags = "400"
Set VerticalGridTimePeriod = mVerticalGridTimePeriod
End Property

Public Property Get VerticalMouseScrollingAllowed() As Boolean
VerticalMouseScrollingAllowed = mVerticalMouseScrollingAllowed
End Property

Public Property Let VerticalMouseScrollingAllowed(ByVal value As Boolean)
mVerticalMouseScrollingAllowed = value
PropertyChanged PropNameVerticalMouseScrollingAllowed
End Property

Public Property Get XAxisRegion() As ChartRegion
Attribute XAxisRegion.VB_MemberFlags = "400"
Set XAxisRegion = mXAxisRegion
End Property

Public Property Get XAxisVisible() As Boolean
Attribute XAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
XAxisVisible = mXAxisVisible
End Property

Public Property Let XAxisVisible(ByVal value As Boolean)
mXAxisVisible = value
sizeRegions
XAxisPicture.Visible = mXAxisVisible
PropertyChanged PropNameXAxisVisible
End Property

Public Property Let XCursorTextStyle(ByVal value As TextStyle)
mXCursorText.LocalStyle = value
End Property

Public Property Get XCursorTextStyle() As TextStyle
Set XCursorTextStyle = mXCursorText.LocalStyle
End Property

Public Property Get YAxisPosition() As Long
Attribute YAxisPosition.VB_MemberFlags = "400"
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisVisible() As Boolean
Attribute YAxisVisible.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisVisible = mYAxisVisible
End Property

Public Property Let YAxisVisible(ByVal value As Boolean)
mYAxisVisible = value
resizeX
PropertyChanged PropNameYAxisVisible
End Property

Public Property Get YAxisWidthCm() As Single
Attribute YAxisWidthCm.VB_ProcData.VB_Invoke_Property = ";Appearance"
YAxisWidthCm = mYAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
If value <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & "YAxisWidthCm", _
            "Y axis Width must be greater than 0"
End If

mYAxisWidthCm = value
resizeX
PropertyChanged PropNameYAxisWidthCm
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddChartRegion(ByVal PercentHeight As Double, _
                    Optional ByVal MinimumPercentHeight As Double, _
                    Optional ByVal Style As ChartRegionStyle, _
                    Optional ByVal yAxisStyle As ChartRegionStyle, _
                    Optional ByVal Name As String) As ChartRegion

Dim ev As CollectionChangeEvent
Dim var As Variant
Dim p As Period
Dim controlIndex As Long
Dim YAxisRegion As ChartRegion
Dim btn As Button


'
' NB: PercentHeight=100 means the region will use whatever space
' is available
'

If Name <> "" Then
    If Not GetChartRegion(Name) Is Nothing Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                "ChartSkil26.Chart::addChartRegion", _
                "Region " & Name & " already exists"
    End If
End If

If Style Is Nothing Then Set Style = getDefaultDataRegionStyle
If yAxisStyle Is Nothing Then Set yAxisStyle = getDefaultYAxisRegionStyle

mRegionsIndex = mRegionsIndex + 1
controlIndex = 1 + (mRegionsIndex - 1) / 2

Set AddChartRegion = New ChartRegion

Load ChartRegionPicture(controlIndex)
ChartRegionPicture(controlIndex).Visible = True
ChartRegionPicture(controlIndex).Align = vbAlignNone
'ChartRegionPicture(controlIndex).Width = _
'    UserControl.ScaleWidth * (mYAxisPosition - ChartLeft) / XAxisPicture.ScaleWidth
ChartRegionPicture(controlIndex).Width = _
    IIf(mYAxisVisible, UserControl.ScaleWidth - mYAxisWidthCm * TwipsPerCm, UserControl.ScaleWidth)
ChartRegionPicture(controlIndex).ZOrder 1

AddChartRegion.Initialise Name, _
                        Me, _
                        createCanvas(ChartRegionPicture(controlIndex))

AddChartRegion.IsDrawingEnabled = (mSuppressDrawingCount > 0)
'addChartRegion.currentTool = mCurrentTool
AddChartRegion.MinimumPercentHeight = MinimumPercentHeight
AddChartRegion.PercentHeight = PercentHeight
AddChartRegion.PointerStyle = mPointerStyle
AddChartRegion.PointerIcon = mPointerIcon
AddChartRegion.PointerCrosshairsColor = mPointerCrosshairsColor
AddChartRegion.PointerDiscColor = mPointerDiscColor
Select Case mPointerMode
Case PointerModeDefault
    AddChartRegion.SetPointerModeDefault
Case PointerModeTool
    AddChartRegion.SetPointerModeTool mToolPointerStyle, mToolIcon
Case PointerModeSelection
    AddChartRegion.SetPointerModeSelection
End Select
AddChartRegion.Left = mScaleLeft
AddChartRegion.RegionNumber = mRegionsIndex
AddChartRegion.Bottom = 0
AddChartRegion.Top = 1
AddChartRegion.PeriodsInView mScaleLeft, mYAxisPosition - 1
AddChartRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
AddChartRegion.SessionStartTime = mSessionStartTime

AddChartRegion.Style = Style

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
If PercentHeight <> 100 Then
    mRegions(mRegionsIndex).PercentHeight = mRegionHeightReductionFactor * PercentHeight
Else
    mRegions(mRegionsIndex).useAvailableSpace = True
End If

Load RegionDividerPicture(controlIndex)
RegionDividerPicture(controlIndex).ZOrder 0

mRegionsIndex = mRegionsIndex + 1
If mRegionsIndex > UBound(mRegions) Then
    ReDim Preserve mRegions(2 * (UBound(mRegions) + 1) - 1) As RegionTableEntry
End If

Set YAxisRegion = New ChartRegion

Load YAxisPicture(controlIndex)
YAxisPicture(controlIndex).Align = vbAlignNone
YAxisPicture(controlIndex).Left = ChartRegionPicture(controlIndex).Width
'YAxisPicture(controlIndex).Width = UserControl.ScaleWidth - YAxisPicture(YAxisPicture.UBound).Left
YAxisPicture(controlIndex).Width = mYAxisWidthCm * TwipsPerCm
YAxisPicture(controlIndex).Visible = mYAxisVisible

YAxisRegion.Initialise "", _
                    Me, _
                    createCanvas(YAxisPicture(controlIndex))

YAxisRegion.RegionNumber = mRegionsIndex
YAxisRegion.IsYAxisRegion = True
YAxisRegion.Style = yAxisStyle
AddChartRegion.YAxisRegion = YAxisRegion

Set mRegions(mRegionsIndex).Region = YAxisRegion

mNumRegionsInUse = mNumRegionsInUse + 1

If sizeRegions Then
    XAxisPicture.Visible = mXAxisVisible
    Set ev.affectedItem = AddChartRegion
    ev.changeType = CollItemAdded
    RaiseEvent RegionsChanged(ev)
Else
    ' can't fit this all in! So remove the added region,
    Set AddChartRegion = Nothing
    Set mRegions(mRegionsIndex).Region = Nothing
    mRegions(mRegionsIndex).PercentHeight = 0
    mRegions(mRegionsIndex).actualHeight = 0
    mRegions(mRegionsIndex).useAvailableSpace = False
    Unload ChartRegionPicture(controlIndex)
    Unload RegionDividerPicture(mRegionsIndex)
    Unload YAxisPicture(controlIndex)
    mRegionsIndex = mRegionsIndex - 2
    mNumRegionsInUse = mNumRegionsInUse - 1
End If

End Function

Public Function addPeriod(ByVal Timestamp As Date) As Period
Set addPeriod = mPeriods.addPeriod(Timestamp)
End Function

Private Function calcScaleLeft() As Single
calcScaleLeft = mYAxisPosition + _
            IIf(mYAxisVisible, mYAxisWidthCm * TwipsPerCm / XAxisPicture.Width * mScaleWidth, 0) - _
            mScaleWidth
End Function

Public Function ClearChart()
Dim i As Long
Dim controlIndex As Long

For i = 1 To mRegionsIndex Step 2
    controlIndex = 1 + (i - 1) / 2
    If Not mRegions(i).Region Is Nothing Then
        mRegions(i).Region.ClearRegion
        Set mRegions(i).Region = Nothing
        mRegions(i + 1).Region.ClearRegion
        Set mRegions(i + 1).Region = Nothing
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
Set mXCursorText = Nothing
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

Initialise
mYAxisPosition = 1
createXAxisRegion
resizeBackground
resizeX
'Resize False

RaiseEvent ChartCleared
Debug.Print "Chart cleared"
End Function

Public Sub DisableDrawing()
SuppressDrawing True
End Sub

Public Sub EnableDrawing()
SuppressDrawing False
End Sub

Public Function GetChartRegion(ByVal Name As String) As ChartRegion
Dim i As Long

Name = UCase$(Name)
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        If UCase$(mRegions(i).Region.Name) = Name Then
            Set GetChartRegion = mRegions(i).Region
            Exit Function
        End If
    End If
Next
                    
End Function

Public Function GetXFromTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal forceNewPeriod As Boolean, _
                Optional ByVal duplicateNumber As Long) As Double
Dim lPeriod As Period
Dim periodEndtime As Date

Select Case BarTimePeriod.units
Case TimePeriodNone, _
        TimePeriodSecond, _
        TimePeriodMinute, _
        TimePeriodHour, _
        TimePeriodDay, _
        TimePeriodWeek, _
        TimePeriodMonth, _
        TimePeriodYear
    
    On Error Resume Next
    Set lPeriod = mPeriods.Item(Timestamp)
    On Error GoTo 0
    
    If lPeriod Is Nothing Then
        If mPeriods.Count = 0 Then
            Set lPeriod = mPeriods.addPeriod(Timestamp)
        ElseIf Timestamp < mPeriods.Item(1).Timestamp Then
            Set lPeriod = mPeriods.Item(1)
            Timestamp = lPeriod.Timestamp
        Else
            Set lPeriod = mPeriods.addPeriod(Timestamp)
        End If
    End If
    
    periodEndtime = BarEndTime(lPeriod.Timestamp, _
                            BarTimePeriod, _
                            SessionStartTime)
    GetXFromTimestamp = lPeriod.PeriodNumber + (Timestamp - lPeriod.Timestamp) / (periodEndtime - lPeriod.Timestamp)
    
Case TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    If Not forceNewPeriod Then
        On Error Resume Next
        Set lPeriod = mPeriods.ItemDup(Timestamp, duplicateNumber)
        On Error GoTo 0
        
        If lPeriod Is Nothing Then
            Set lPeriod = mPeriods.addPeriod(Timestamp, True)
        End If
        GetXFromTimestamp = lPeriod.PeriodNumber
    Else
        Set lPeriod = mPeriods.addPeriod(Timestamp, True)
        GetXFromTimestamp = lPeriod.PeriodNumber
    End If
End Select

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

Public Function IsTimeInSession(ByVal Timestamp As Date) As Boolean

If Timestamp >= mCurrentSessionStartTime And _
    Timestamp < mCurrentSessionEndTime _
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

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "RemoveChartRegion" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub ScrollX(ByVal value As Long)
Dim Region As ChartRegion
Dim i As Long
Dim failpoint As Long
On Error GoTo Err

If value = 0 Then Exit Sub

If (LastVisiblePeriod + value) > _
        (mPeriods.CurrentPeriodNumber + ChartWidth - 1) Then
    value = mPeriods.CurrentPeriodNumber + ChartWidth - 1 - LastVisiblePeriod
ElseIf (LastVisiblePeriod + value) < 1 Then
    value = 1 - LastVisiblePeriod
End If

mYAxisPosition = mYAxisPosition + value
mScaleLeft = calcScaleLeft
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
mXAxisRegion.PeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
setHorizontalScrollBar

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "ScrollX" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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

RaiseEvent PointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPointerModeDefault" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub SetPointerModeSelection()
Dim i As Long
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err

mPointerMode = PointerModeSelection

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.SetPointerModeSelection
    End If
Next

RaiseEvent PointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPointerModeSelection" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
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
mToolPointerStyle = toolPointerStyle
Set mToolIcon = icon

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

RaiseEvent PointerModeChanged
Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "SetPointerModeTool" & "." & failpoint & IIf(Err.Source <> "", vbCrLf & Err.Source, "")
Dim errDescription As String: errDescription = Err.Description
gErrorLogger.Log LogLevelSevere, "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Public Sub ShowGrid()
Dim i As Long
Dim Region As ChartRegion

If Not mHideGrid Then Exit Sub

mHideGrid = False
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        Region.ShowGrid
    End If
Next
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcAvailableHeight() As Long
calcAvailableHeight = IIf(mXAxisVisible, XAxisPicture.Top, UserControl.ScaleHeight) - _
                    mNumRegionsInUse * RegionDividerPicture(0).Height - _
                    IIf(mToolbarVisible, Toolbar1.Height, 0)
If calcAvailableHeight < 0 Then calcAvailableHeight = 0
End Function

Private Sub CalcSessionTimes(ByVal Timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date)
Dim i As Long

i = -1
Do
    i = i + 1
Loop Until calcSessionTimesHelper(Timestamp + i, SessionStartTime, SessionEndTime)
End Sub

Friend Function calcSessionTimesHelper(ByVal Timestamp As Date, _
                            ByRef SessionStartTime As Date, _
                            ByRef SessionEndTime As Date) As Boolean
Dim referenceDate As Date
Dim referenceTime As Date
Dim weekday As Long

referenceDate = DateValue(Timestamp)
referenceTime = TimeValue(Timestamp)

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
                ByVal Surface As PictureBox) As Canvas
Set createCanvas = New Canvas
createCanvas.Surface = Surface
End Function

Private Sub createXAxisRegion()
Dim aFont As StdFont
Dim lCanvas As Canvas

Set mXAxisRegion = New ChartRegion

'Set mRegions(0).Region = mXAxisRegion
Set lCanvas = createCanvas(XAxisPicture)
mXAxisRegion.Initialise "", _
                    Me, _
                    lCanvas
                        
mXAxisRegion.IsXAxisRegion = True
mXAxisRegion.VerticalGridTimePeriod = mVerticalGridTimePeriod
mXAxisRegion.Bottom = 0
mXAxisRegion.Top = 1
mXAxisRegion.SessionStartTime = mSessionStartTime
mXAxisRegion.HasGrid = False
mXAxisRegion.HasGridText = True

Set mXCursorText = mXAxisRegion.AddText(LayerNumbers.LayerPointer)
mXCursorText.Align = AlignTopCentre

Dim txtStyle As New TextStyle
txtStyle.Color = vbBlack
txtStyle.Box = True
txtStyle.BoxFillColor = vbWhite
txtStyle.BoxStyle = LineSolid
txtStyle.BoxColor = vbBlack
Set aFont = New StdFont
aFont.Name = "Arial"
aFont.size = 8
aFont.Underline = False
aFont.Bold = False
txtStyle.Font = aFont
mXCursorText.LocalStyle = txtStyle
End Sub

Private Sub displayXAxisLabel(ByVal X As Single, ByVal Y As Single)
Dim thisPeriod As Period
Dim PeriodNumber As Long
Dim prevPeriodNumber As Long
Dim prevPeriod As Period

If mXAxisRegion Is Nothing Then createXAxisRegion

If Round(X) >= mYAxisPosition Then Exit Sub
If mPeriods.Count = 0 Then Exit Sub

On Error Resume Next
PeriodNumber = Round(X)
Set thisPeriod = mPeriods(PeriodNumber)
On Error GoTo 0
If thisPeriod Is Nothing Then
    mXCursorText.Text = ""
    Exit Sub
End If

mXCursorText.position = mXAxisRegion.NewPoint( _
                            PeriodNumber, _
                            0, _
                            CoordsLogical, _
                            CoordsCounterDistance)

Select Case mBarTimePeriod.units
Case TimePeriodNone, TimePeriodMinute, TimePeriodHour
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.Timestamp, vbShortTime)
Case TimePeriodSecond, TimePeriodVolume, TimePeriodTickVolume, TimePeriodTickMovement
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.Timestamp, vbLongTime)
Case Else
    mXCursorText.Text = FormatDateTime(thisPeriod.Timestamp, vbShortDate)
End Select

End Sub

Private Function getDefaultDataRegionStyle() As ChartRegionStyle
Static defaultDataRegionStyle As ChartRegionStyle
If defaultDataRegionStyle Is Nothing Then Set defaultDataRegionStyle = New ChartRegionStyle
Set getDefaultDataRegionStyle = defaultDataRegionStyle
End Function

Private Function getDefaultYAxisRegionStyle() As ChartRegionStyle
Static defaultYAxisRegionStyle As ChartRegionStyle
If defaultYAxisRegionStyle Is Nothing Then Set defaultYAxisRegionStyle = New ChartRegionStyle
Set getDefaultYAxisRegionStyle = defaultYAxisRegionStyle
End Function

Private Sub Initialise()
Static firstInitialisationDone As Boolean
Dim i As Long
Dim btn As Button

For Each btn In Toolbar1.Buttons
    btn.value = tbrUnpressed
    btn.Enabled = False
Next

mPrevHeight = UserControl.Height

ReDim mRegions(3) As RegionTableEntry
mRegionsIndex = 0
mNumRegionsInUse = 0
mRegionHeightReductionFactor = 1

Set mPeriods = New Periods
mPeriods.Chart = Me

mBarTimePeriodSet = False

If Not firstInitialisationDone Then
    
    firstInitialisationDone = True
    
    ' these values are only set once when the control initialises
    ' if the chart is subsequently cleared, any values set by the
    ' application remain in force
    
    mAutoscrolling = PropDfltAutoscrolling
    Set mBarTimePeriod = GetTimePeriod(PropDfltPeriodLength, PropDfltPeriodUnits)
    mPointerCrosshairsColor = PropDfltPointerCrosshairsColor
    mPointerDiscColor = PropDfltPointerDiscColor
    mHorizontalScrollBarVisible = PropDfltHorizontalScrollBarVisible
    mToolbarVisible = PropDfltToolbarVisible
    'HScroll.Height = HorizScrollBarHeight
    HScroll.Visible = mHorizontalScrollBarVisible
    Set mVerticalGridTimePeriod = GetTimePeriod(PropDfltVerticalGridSpacing, PropDfltVerticalGridUnits)
    mVerticalGridTimePeriodSet = False
    
    mTwipsPerBar = PropDfltTwipsPerBar
    
    mXAxisVisible = PropDfltXAxisVisible
    mYAxisWidthCm = PropDfltYAxisWidthCm
    mYAxisVisible = PropDfltYAxisVisible

    mHorizontalMouseScrollingAllowed = PropDfltHorizontalMouseScrollingAllowed
    mVerticalMouseScrollingAllowed = PropDfltVerticalMouseScrollingAllowed

End If

mPointerMode = PointerModes.PointerModeDefault

mYAxisPosition = 1
mScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar) - 0.5!
mScaleLeft = calcScaleLeft
mScaleHeight = -100
mScaleTop = 100

HScroll.value = 0

ChartRegionPicture(0).Visible = True
resizeBackground

End Sub

Private Sub mouseMove( _
                ByVal index As Long, _
                ByVal Button As Long, _
                ByVal Shift As Long, _
                ByRef X As Single, _
                ByRef Y As Single)
Dim i As Long
Dim Region As ChartRegion

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set Region = mRegions(i).Region
        If i = (2 * index - 1) Then
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & Y
            If (mPointerMode = PointerModeDefault And _
                    ((Region.CursorSnapsToTickBoundaries And Not CBool(Shift And vbCtrlMask)) Or _
                    (Not Region.CursorSnapsToTickBoundaries And CBool(Shift And vbCtrlMask)))) Or _
                (mPointerMode = PointerModeTool And CBool(Shift And vbCtrlMask)) _
            Then
                Dim YScaleQuantum As Double
                YScaleQuantum = Region.YScaleQuantum
                If YScaleQuantum <> 0 Then Y = YScaleQuantum * Int((Y + YScaleQuantum / 10000) / YScaleQuantum)
            End If
            Region.DrawCursor Button, Shift, X, Y
            
        Else
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & X & " y=" & MinusInfinitySingle
            Region.DrawCursor Button, Shift, X, MinusInfinitySingle
        End If
    End If
Next
displayXAxisLabel Round(X), 100
End Sub

Private Sub mouseScroll( _
                ByVal index As Long, _
                ByVal Button As Long, _
                ByVal Shift As Long, _
                ByRef X As Single, _
                ByRef Y As Single)

If mHorizontalMouseScrollingAllowed Then
    ' the chart needs to be scrolled so that current mouse position
    ' is the value contained in mLeftDragStartPosnX
    If mLeftDragStartPosnX <> Int(X) Then
        If (LastVisiblePeriod + mLeftDragStartPosnX - Int(X)) <= _
                (mPeriods.CurrentPeriodNumber + ChartWidth - 1) And _
            (LastVisiblePeriod + mLeftDragStartPosnX - Int(X)) >= 1 _
        Then
            ScrollX mLeftDragStartPosnX - Int(X)
        End If
    End If
End If
If mVerticalMouseScrollingAllowed Then
    If mLeftDragStartPosnY <> Y Then
        With mRegions(2 * index - 1).Region
            If Not .Autoscaling Then
                .ScrollVertical mLeftDragStartPosnY - Y
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

'gLogger.Log LogLevelDetail, "ChartSkil: Resize: enter"

resizeBackground

If resizeWidth Then
    HScroll.Width = UserControl.Width
    XAxisPicture.Width = UserControl.Width
    Toolbar1.Width = UserControl.Width
    resizeX
End If

failpoint = 200

If resizeHeight Then
    HScroll.Top = UserControl.Height - IIf(mHorizontalScrollBarVisible, HScroll.Height, 0)
    XAxisPicture.Top = HScroll.Top - IIf(mXAxisVisible, XAxisPicture.Height, 0)
    sizeRegions
End If
'paintAll

'gLogger.Log LogLevelDetail, "ChartSkil: Resize: exit"

Exit Sub

Err:
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "Resize" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "Resize" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Sub

Private Sub resizeBackground()
If mNumRegionsInUse > 0 Then Exit Sub
XAxisPicture.Visible = False
ChartRegionPicture(0).Visible = False
ChartRegionPicture(0).Move 0, 0, UserControl.Width, UserControl.Height
mBackGroundCanvas.GradientFillColors = mChartBackGradientFillColors
mBackGroundCanvas.Left = 0
mBackGroundCanvas.Right = 1
mBackGroundCanvas.Bottom = 0
mBackGroundCanvas.Top = 1
mBackGroundCanvas.PaintBackground
mBackGroundCanvas.ZOrder 1
ChartRegionPicture(0).Visible = True
End Sub

Private Sub resizeX()
Dim i As Long
Dim Region As ChartRegion

Dim failpoint As Long
On Error GoTo Err


failpoint = 100

'If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":resizeX Enter"


failpoint = 200

'mScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar) - 0.5!
mScaleWidth = CSng(XAxisPicture.Width) / CSng(mTwipsPerBar)
mScaleLeft = calcScaleLeft


failpoint = 400

For i = 1 To ChartRegionPicture.UBound
    If (UserControl.Width - YAxisPicture(i).Width) > 0 Then
        YAxisPicture(i).Left = UserControl.Width - IIf(mYAxisVisible, YAxisPicture(i).Width, 0)
        ChartRegionPicture(i).Width = YAxisPicture(i).Left
    End If
Next


failpoint = 500

For i = 0 To RegionDividerPicture.UBound
    RegionDividerPicture(i).Width = UserControl.Width
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
    mXAxisRegion.PeriodsInView mScaleLeft, mScaleLeft + mScaleWidth - 1
End If


failpoint = 800

setHorizontalScrollBar

'If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":resizeX Exit"

Exit Sub

Err:
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "resizeX" & "." & failpoint & _
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

If mPeriods.CurrentPeriodNumber + ChartWidth - 1 > 32767 Then

    failpoint = 100

    HScroll.Max = 32767
ElseIf mPeriods.CurrentPeriodNumber + ChartWidth - 1 < 1 Then

    failpoint = 200

    HScroll.Max = 1
Else

    failpoint = 300
    
    HScroll.Max = mPeriods.CurrentPeriodNumber + ChartWidth - 1
End If
HScroll.Min = 0


failpoint = 400

' NB the following calculation has to be done using doubles as for very large charts it can cause an overflow using integers
hscrollVal = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
If hscrollVal > HScroll.Max Then
    HScroll.value = HScroll.Max
ElseIf hscrollVal < HScroll.Min Then
    HScroll.value = HScroll.Min
Else
    HScroll.value = Round(CDbl(HScroll.Max) * CDbl(LastVisiblePeriod) / CDbl((mPeriods.CurrentPeriodNumber + ChartWidth - 1)))
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
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "setHorizontalScrollBar" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            errDescription
Err.Raise errNumber, _
        ProjectName & "." & ModuleName & ":" & "mTimer_TimerExpired" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        errDescription

End Sub

Private Sub setSession( _
                ByVal Timestamp As Date)
If Timestamp >= mCurrentSessionEndTime Or _
    Timestamp < mReferenceTime _
Then
    mReferenceTime = Timestamp
    CalcSessionTimes Timestamp, mCurrentSessionStartTime, mCurrentSessionEndTime
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
' NB: PercentHeight=100 means the region will use whatever space
' is available
'
Dim i As Long
Dim Top As Long
Dim aRegion As ChartRegion
Dim numAvailableSpaceRegions As Long
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim availableHeight As Long     ' the space available for the region picture boxes
                                ' excluding the divider pictures
Dim numRegionsSized As Long
Dim HeightReductionFactor As Double
Dim failpoint As Long
On Error GoTo Err

'If gLogger.isLoggable(LogLevelHighDetail) Then gLogger.Log LogLevelHighDetail, ProjectName & "." & ModuleName & ":sizeRegions Enter"


failpoint = 100

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
'        mRegions(i).PercentHeight = aRegion.PercentHeight
        If Not mRegions(i).useAvailableSpace Then
            availableSpacePercent = availableSpacePercent - mRegions(i).PercentHeight
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).PercentHeight
        Else
            If aRegion.MinimumPercentHeight <> 0 Then
                availableSpacePercent = availableSpacePercent - aRegion.MinimumPercentHeight
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

HeightReductionFactor = 1
Do While availableSpacePercent < 0
    availableSpacePercent = 100
    nonFixedAvailableSpacePercent = 100
    mRegionHeightReductionFactor = mRegionHeightReductionFactor * 0.95
    HeightReductionFactor = HeightReductionFactor * 0.95
    For i = 1 To mRegionsIndex Step 2
        If Not mRegions(i).Region Is Nothing Then
            Set aRegion = mRegions(i).Region
            If Not mRegions(i).useAvailableSpace Then
                If aRegion.MinimumPercentHeight <> 0 Then
                    If mRegions(i).PercentHeight * mRegionHeightReductionFactor >= _
                        aRegion.MinimumPercentHeight _
                    Then
                        mRegions(i).PercentHeight = mRegions(i).PercentHeight * mRegionHeightReductionFactor
                    Else
                        mRegions(i).PercentHeight = aRegion.MinimumPercentHeight
                    End If
                    totalMinimumPercents = totalMinimumPercents + aRegion.MinimumPercentHeight
                Else
                    mRegions(i).PercentHeight = mRegions(i).PercentHeight * mRegionHeightReductionFactor
                End If
                availableSpacePercent = availableSpacePercent - mRegions(i).PercentHeight
                nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).PercentHeight
            Else
                If aRegion.MinimumPercentHeight <> 0 Then
                    availableSpacePercent = availableSpacePercent - aRegion.MinimumPercentHeight
                    totalMinimumPercents = totalMinimumPercents + aRegion.MinimumPercentHeight
                End If
            End If
        End If
    Next
    If totalMinimumPercents > 100 Then
        ' can't possibly fit this all in!
        sizeRegions = False
        'If gLogger.isLoggable(LogLevelMediumDetail) Then gLogger.Log LogLevelMediumDetail, ProjectName & "." & ModuleName & ":sizeRegions Exit"
        Exit Function
    End If
Loop


failpoint = 300

If numAvailableSpaceRegions = 0 Then
    ' we must adjust the percentages on the other regions so they
    ' total 100.
    For i = 1 To mRegionsIndex Step 2
        mRegions(i).PercentHeight = 100 * mRegions(i).PercentHeight / (100 - nonFixedAvailableSpacePercent)
    Next
End If

' calculate the actual available Height to put these regions in
availableHeight = calcAvailableHeight

' first set Heights for fixed Height regions

failpoint = 400

For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).useAvailableSpace Then
        mRegions(i).actualHeight = mRegions(i).PercentHeight * availableHeight / 100
        Debug.Assert mRegions(i).actualHeight >= 0
    End If
Next


failpoint = 500

' now set Heights for 'available space' regions with a minimum Height
' that needs to be respected
For i = 1 To mRegionsIndex Step 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If mRegions(i).useAvailableSpace Then
            mRegions(i).actualHeight = 0
            If aRegion.MinimumPercentHeight <> 0 Then
                If (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) < aRegion.MinimumPercentHeight Then
                    mRegions(i).actualHeight = aRegion.MinimumPercentHeight * availableHeight / 100
                    Debug.Assert mRegions(i).actualHeight >= 0
                    nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.MinimumPercentHeight
                    numAvailableSpaceRegions = numAvailableSpaceRegions - 1
                End If
            End If
        End If
    End If
Next


failpoint = 600

' finally set Heights for all other 'available space' regions
For i = 1 To mRegionsIndex Step 2
    If mRegions(i).useAvailableSpace And _
        mRegions(i).actualHeight = 0 _
    Then
        mRegions(i).actualHeight = (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) * availableHeight / 100
        Debug.Assert mRegions(i).actualHeight >= 0
    End If
Next


failpoint = 700

' Now actually set the Heights and positions for the picture boxes

Top = IIf(mToolbarVisible, Toolbar1.Height, 0)

Dim controlIndex As Long
    
For i = 1 To mRegionsIndex Step 2
    controlIndex = 1 + (i - 1) / 2
    If Not mRegions(i).Region Is Nothing Then
        Set aRegion = mRegions(i).Region
        If Not IsDrawingEnabled Then
            ChartRegionPicture(controlIndex).Height = mRegions(i).actualHeight
            YAxisPicture(controlIndex).Height = mRegions(i).actualHeight
            ChartRegionPicture(controlIndex).Top = Top
            YAxisPicture(controlIndex).Top = Top
            aRegion.resizedY
        End If
        Top = Top + mRegions(i).actualHeight
        numRegionsSized = numRegionsSized + 1
        If Not IsDrawingEnabled Then
            RegionDividerPicture(controlIndex).Top = Top
            RegionDividerPicture(controlIndex).Visible = True
        End If
        If numRegionsSized <> mNumRegionsInUse Then
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbSizeNS
        Else
            RegionDividerPicture(controlIndex).MousePointer = MousePointerConstants.vbDefault
        End If
        Top = Top + RegionDividerPicture(controlIndex).Height
    Else
        If Not IsDrawingEnabled Then
            ChartRegionPicture(controlIndex).Visible = False
            YAxisPicture(controlIndex).Visible = False
            RegionDividerPicture(controlIndex).Visible = False
        End If
    End If
Next

sizeRegions = True

'If gLogger.isLoggable(LogLevelHighDetail) Then gLogger.Log LogLevelHighDetail, ProjectName & "." & ModuleName & ":sizeRegions Exit"

Exit Function

Err:
gErrorLogger.Log LogLevelSevere, "Error at: " & ProjectName & "." & ModuleName & ":" & "sizeRegions" & "." & failpoint & _
                            IIf(Err.Source <> "", vbCrLf & Err.Source, "") & vbCrLf & _
                            Err.Description
Err.Raise Err.Number, _
        ProjectName & "." & ModuleName & ":" & "sizeRegions" & "." & failpoint & _
        IIf(Err.Source <> "", vbCrLf & Err.Source, ""), _
        Err.Description

End Function

Private Sub SuppressDrawing(ByVal suppress As Boolean)
Dim i As Long
Dim Region As ChartRegion
If suppress Then
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
        Region.IsDrawingEnabled = (mSuppressDrawingCount > 0)
    End If
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.IsDrawingEnabled = (mSuppressDrawingCount > 0)
End Sub

Public Property Get TwipsPerBar() As Long
TwipsPerBar = mTwipsPerBar
End Property

Private Sub zoom(ByRef rect As TRectangle)

End Sub

