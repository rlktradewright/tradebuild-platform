VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Chart 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7575
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10665
   ScaleHeight     =   7575
   ScaleWidth      =   10665
   Begin MSComCtl2.FlatScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   7320
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   2
      Arrows          =   65536
      Orientation     =   1245185
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   1800
      Top             =   2040
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
      Left            =   480
      Top             =   2040
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
      TabIndex        =   3
      Top             =   6240
      Visible         =   0   'False
      Width           =   9375
   End
   Begin VB.PictureBox YAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Index           =   0
      Left            =   8400
      ScaleHeight     =   615
      ScaleWidth      =   975
      TabIndex        =   4
      Top             =   6360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.PictureBox XAxisPicture 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
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
      BackColor       =   &H80000005&
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
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ChartArea.ctx":9F4C
            Key             =   ""
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
      Top             =   960
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
      TabIndex        =   5
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
            Key             =   "showcursor"
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
' Member variables and constants
'================================================================================

Private Const DefaultTwipsPerBar As Long = 150

Private mRegions() As RegionTableEntry
Private mRegionsIndex As Long

Private WithEvents mPeriods As periods
Attribute mPeriods.VB_VarHelpID = -1

Private mAutoscale As Boolean
Private mScaleWidth As Single
Private mScaleHeight As Single
Private mScaleLeft As Single
Private mScaleTop As Single

Private mPrevHeight As Single

Private mTwipsPerBar As Long

Private mXAxisRegion As ChartRegion
Private mXCursorText As text

Private mYAxisPosition As Long
Private mYAxisWidthCm As Single

Private mSessionStartTime As Date
Private mPeriodLengthMinutes As Long

Private mVerticalGridSpacing As Long
Private mVerticalGridUnits As TimeUnits

Private mBackColour As Long
Private mGridColour As Long
Private mGridTextColor As Long
Private mShowGrid As Boolean
Private mShowCrosshairs As Boolean

Private mNotFirstMouseMove As Boolean
Private mPrevCursorX As Single
Private mPrevCursorY As Single

Private mSuppressDrawing As Boolean
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

'================================================================================
' Enums
'================================================================================

Enum ArrowStyles
    ArrowNone
    ArrowSingleOpen
    ArrowDoubleOpen
    ArrowClosed
    ArrowSingleBar
    ArrowDoubleBar
    ArrowLollipop
    ArrowDiamond
    ArrowBarb
End Enum

Public Enum CoordinateSystems
    CoordsLogical = 0
    CoordsRelative
    CoordsDistance        ' Measured from left or bottom of region
    CoordsCounterDistance ' Measured from right or top of region
End Enum

Enum DrawModes
    DrawModeBlackness = vbBlackness
    DrawModeCopyPen = vbCopyPen
    DrawModeInvert = vbInvert
    DrawModeMaskNotPen = vbMaskNotPen
    DrawModeMaskPen = vbMaskPen
    DrawModeMaskPenNot = vbMaskPenNot
    DrawModeMergeNotPen = vbMergeNotPen
    DrawModeMergePen = vbMergePen
    DrawModeMergePenNot = vbMergePenNot
    DrawModeNop = vbNop
    DrawModeNotCopyPen = vbNotCopyPen
    DrawModeNotMaskPen = vbNotMaskPen
    DrawModeNotMergePen = vbNotMergePen
    DrawModeNotXorPen = vbNotXorPen
    DrawModeWhiteness = vbWhiteness
    DrawModeXorPen = vbXorPen
End Enum

Enum ErrorCodes
    InvalidPropertyValue = 380
End Enum

Enum FillStyles
    FillSolid = vbFSSolid ' 0 Solid
    FillTransparent = vbFSTransparent ' 1 (Default) Transparent
    FillHorizontalLine = vbHorizontalLine ' 2 Horizontal Line
    FillVerticalLine = vbVerticalLine ' 3 Vertical Line
    FillUpwardDiagonal = vbUpwardDiagonal ' 4 Upward Diagonal
    FillDownwardDiagonal = vbDownwardDiagonal ' 5 Downward Diagonal
    FillCross = vbCross ' 6 Cross
    FillDiagonalCross = vbDiagonalCross ' 7 Diagonal Cross
End Enum

Enum LineStyles
    LineSolid = vbSolid
    LineDash = vbDash
    LineDot = vbDot
    LineDashDot = vbDashDot
    LineDashDotDot = vbDashDotDot
    LineInvisible = vbInvisible
    LineInsideSolid = vbInsideSolid
End Enum

Enum PointerStyles
    PointerNone
    PointerCrosshairs
    PointerDisc
End Enum

Enum TextAlignModes
    AlignTopLeft
    AlignCentreLeft
    AlignBottomLeft
    AlignTopCentre
    AlignCentreCentre
    AlignBottomCentre
    AlignTopRight
    AlignCentreRight
    AlignBottomRight
End Enum

Enum TimeUnits
    TimeSecond
    TimeMinute
    TimeHour
    TimeDay
    TimeWeek
    TimeMonth
    TimeYear
End Enum

Enum ToolTypes
    ToolPointer
    ToolLine
    ToolLineExtended
    ToolLineRay
    ToolLineHorizontal
    ToolLineVertical
    ToolFibonacciRetracement
    ToolFibonacciExtension
    ToolFibonacciCircle
    ToolFibonacciTime
    ToolRegressionChannel
    ToolRegressionEnvelope
    ToolText
    ToolPitchfork
    ToolCircle
    ToolRectangle
End Enum

Enum Verticals
    VerticalNot
    VerticalUp
    VerticalDown
End Enum

Enum Quadrants
    NE
    NW
    SW
    SE
End Enum

'================================================================================
' User Control Event Handlers
'================================================================================

Private Sub UserControl_Initialize()
initialise
createXAxisRegion
End Sub

Private Sub UserControl_Paint()
Static paintcount As Long
paintcount = paintcount + 1
Debug.Print "Control_paint" & paintcount
mPainted = True
paintAll
End Sub

Private Sub UserControl_Resize()
Static resizeCount As Long
resizeCount = resizeCount + 1
'debug.print "Control_resize: count = " & resizeCount
mNotFirstMouseMove = False
HScroll.top = UserControl.height - HScroll.height
HScroll.width = UserControl.width
XAxisPicture.top = HScroll.top - XAxisPicture.height
XAxisPicture.width = UserControl.width
Toolbar1.width = UserControl.width
resizeX
resizeY
paintAll
'debug.print "Exit Control_resize"
End Sub

Private Sub UserControl_Terminate()
Debug.Print "ChartSkil Usercontrol terminated"
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "autoscale", mAutoscale, True
End Sub

'================================================================================
' ChartRegionPicture Event Handlers
'================================================================================

Private Sub ChartRegionPicture_MouseDown( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(x)
mLeftDragStartPosnY = y
End Sub

Private Sub ChartRegionPicture_MouseMove(index As Integer, _
                                Button As Integer, _
                                Shift As Integer, _
                                x As Single, _
                                y As Single)

Dim region As ChartRegion
Dim i As Long

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
            With mRegions(index - 1).region
                If Not .autoscale Then
                    .scrollVertical mLeftDragStartPosnY - y
                End If
            End With
        End If
    End If
Else
    For i = 0 To mRegionsIndex
        Set region = mRegions(i).region
        If i = index - 1 Then
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & y
            region.MouseMove Button, Shift, x, y
        Else
            'debug.print "Mousemove: index=" & index & " region=" & i & " x=" & x & " y=" & MinusInfinitySingle
            region.MouseMove Button, Shift, x, MinusInfinitySingle
        End If
    Next
    displayXAxisLabel x, 100
End If
End Sub

Private Sub ChartRegionPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
If Button = vbLeftButton Then mLeftDragging = False
End Sub

'================================================================================
' HScroll Event Handlers
'================================================================================

Private Sub HScroll_Change()
lastVisiblePeriod = Round((CLng(HScroll.value) - CLng(HScroll.Min)) / (CLng(HScroll.Max) - CLng(HScroll.Min)) * (mPeriods.currentPeriodNumber + chartWidth - 1))
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
If Button = vbLeftButton Then mLeftDragging = True
mLeftDragStartPosnX = Int(x)
mLeftDragStartPosnY = y
mUserResizingRegions = True
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

If Not mLeftDragging = True Then Exit Sub

currRegion = index  ' we resize the region below the divider
vertChange = mLeftDragStartPosnY - y
newHeight = mRegions(currRegion).actualHeight + vertChange

' the region table indicates the requested percentage used by each region
' and the actual height allocation. We need to work out the new percentage
' for the region to be resized.

prevPercentHeight = mRegions(currRegion).region.percentheight
If Not mRegions(currRegion).useAvailableSpace Then
    mRegions(currRegion).region.percentheight = mRegions(currRegion).percentheight * newHeight / mRegions(currRegion).actualHeight
Else
    ' this is a 'use available space' region that's being resized. Now change
    ' it to use a specific percentage
    mRegions(currRegion).region.percentheight = 100 * newHeight / calcAvailableHeight
End If

If sizeRegions Then
    paintAll
Else
    ' the regions couldn't be resized so reset the region's percent height
    mRegions(currRegion).region.percentheight = prevPercentHeight
End If
End Sub

Private Sub RegionDividerPicture_MouseUp( _
                            index As Integer, _
                            Button As Integer, _
                            Shift As Integer, _
                            x As Single, _
                            y As Single)
If Button = vbLeftButton Then mLeftDragging = False
mUserResizingRegions = False
End Sub

'================================================================================
' Toolbar1 Event Handlers
'================================================================================

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.key
Case "showcrosshair"
    showCrosshairs = True
Case "showcursor"
    showCrosshairs = False
Case "reducespacing"
    If twipsPerBar >= 50 Then
        twipsPerBar = twipsPerBar - 25
    End If
    If twipsPerBar < 50 Then
        Button.Enabled = False
    End If
Case "increasespacing"
    twipsPerBar = twipsPerBar + 25
    Toolbar1.Buttons("reducespacing").Enabled = True
Case "scrollleft"
    scrollX -(chartWidth * 0.2)
Case "scrollright"
    scrollX chartWidth * 0.2
Case "scrollend"
    lastVisiblePeriod = currentPeriodNumber
End Select

End Sub

'================================================================================
' mPeriods Event Handlers
'================================================================================

Private Sub mPeriods_PeriodAdded(ByVal period As period)
Dim i As Long
Dim region As ChartRegion

period.backColor = mBackColour

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.addperiod period.periodNumber, period.timestamp
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.addperiod period.periodNumber, period.timestamp
setHorizontalScrollBar
End Sub

'================================================================================
' Properties
'================================================================================

Public Property Get allowHorizontalMouseScrolling() As Boolean
allowHorizontalMouseScrolling = mAllowHorizontalMouseScrolling
End Property

Public Property Let allowHorizontalMouseScrolling(ByVal value As Boolean)
mAllowHorizontalMouseScrolling = value
PropertyChanged "allowHorizontalMouseScrolling"
End Property

Public Property Get allowVerticalMouseScrolling() As Boolean
allowVerticalMouseScrolling = mAllowVerticalMouseScrolling
End Property

Public Property Let allowVerticalMouseScrolling(ByVal value As Boolean)
mAllowVerticalMouseScrolling = value
PropertyChanged "allowVerticalMouseScrolling"
End Property

Public Property Get autoscale() As Boolean
autoscale = mAutoscale
End Property

Public Property Let autoscale(ByVal value As Boolean)
mAutoscale = value
PropertyChanged "autoscale"
End Property

'Public Property Get barSpacingPercent() As Single
'barSpacingPercent = mBarSpacingPercent
'End Property
'
'Public Property Let barSpacingPercent(ByVal value As Single)
'mBarSpacingPercent = value
'mCandleWidth = 100! / (100! + mBarSpacingPercent)
'PropertyChanged "barspacingpercent"
'End Property

Public Property Get chartBackColor() As Long
chartBackColor = mBackColour
End Property

Public Property Let chartBackColor(ByVal val As Long)
Dim i As Long

mBackColour = val
XAxisPicture.backColor = val

For i = 0 To mRegionsIndex
    mRegions(i).region.regionBackColor = val
Next
paintAll
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

Public Property Get firstVisiblePeriod() As Long
firstVisiblePeriod = mScaleLeft
End Property

Public Property Let firstVisiblePeriod(ByVal value As Long)
scrollX value - mScaleLeft + 1
End Property

Public Property Get gridColor() As Long
gridColor = mGridColour
End Property

Public Property Let gridColor(ByVal val As Long)
mGridColour = val
End Property

Public Property Get gridTextColor() As Long
gridTextColor = mGridTextColor
End Property

Public Property Let gridTextColor(ByVal val As Long)
mGridTextColor = val
End Property

Public Property Get lastVisiblePeriod() As Long
lastVisiblePeriod = mYAxisPosition - 1
End Property

Public Property Let lastVisiblePeriod(ByVal value As Long)
scrollX value - mYAxisPosition + 1
End Property

Public Property Get periodLengthMinutes() As Long
periodLengthMinutes = mPeriodLengthMinutes
End Property

Public Property Let periodLengthMinutes(ByVal val As Long)
mPeriodLengthMinutes = val
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.periodLengthMinutes = val
End Property

Public Property Get periods() As periods
Set periods = mPeriods
End Property

Public Property Get sessionStartTime() As Date
sessionStartTime = mSessionStartTime
End Property

Public Property Let sessionStartTime(ByVal val As Date)
If CDbl(val) >= 1 Then _
    Err.Raise CommonErrorCodes.InvalidPropertyValue, _
                "ChartSkil.Chart::(Let)sessionStartTime", _
                "Value must be a time only"
mSessionStartTime = val
End Property

Public Property Get showCrosshairs() As Boolean
showCrosshairs = mShowCrosshairs
End Property

Public Property Let showCrosshairs(ByVal val As Boolean)
Dim i As Long
Dim region As ChartRegion
mShowCrosshairs = val
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    If val Then
        region.pointerStyle = PointerCrosshairs
    Else
        region.pointerStyle = PointerDisc
    End If
Next
End Property

Public Property Get showGrid() As Boolean
showGrid = mShowGrid
End Property

Public Property Let showGrid(ByVal val As Boolean)
mShowGrid = val
End Property

Public Property Get showHorizontalScrollBar() As Boolean
showHorizontalScrollBar = mShowHorizontalScrollBar
End Property

Public Property Let showHorizontalScrollBar(ByVal val As Boolean)
mShowHorizontalScrollBar = val
If mShowHorizontalScrollBar Then
    HScroll.height = 255
Else
    HScroll.height = 0
End If
resizeY
End Property

Public Property Get suppressDrawing() As Boolean
suppressDrawing = mSuppressDrawing
End Property

Public Property Let suppressDrawing(ByVal val As Boolean)
Dim i As Long
Dim region As ChartRegion
mSuppressDrawing = val
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.suppressDrawing = val
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.suppressDrawing = val
End Property

Public Property Get twipsPerBar() As Long
twipsPerBar = mTwipsPerBar
End Property

Public Property Let twipsPerBar(ByVal val As Long)
mTwipsPerBar = val
resizeX
setHorizontalScrollBar
paintAll
End Property

Public Property Let verticalGridSpacing(ByVal value As Long)
If value < 0 Then _
    Err.Raise CommonErrorCodes.InvalidPropertyValue, _
                "ChartSkil.Chart::(Let)verticalGridSpacing", _
                "Value must be >= 0"
mVerticalGridSpacing = value
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.verticalGridSpacing = mVerticalGridSpacing
End Property

Public Property Get verticalGridSpacing() As Long
verticalGridSpacing = mVerticalGridSpacing
End Property

Public Property Let verticalGridUnits(ByVal value As TimeUnits)
Select Case value
Case TimeSecond
Case TimeMinute
Case TimeHour
Case TimeDay
Case TimeWeek
Case TimeMonth
Case TimeYear
Case Else
    Err.Raise CommonErrorCodes.InvalidPropertyValue, _
                "ChartSkil.Chart::(Let)verticalGridUnits", _
                "Value must be a member of the TimeUnits enum"
End Select
mVerticalGridUnits = value
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.verticalGridUnits = mVerticalGridUnits
End Property

Public Property Get verticalGridUnits() As TimeUnits
verticalGridUnits = mVerticalGridUnits
End Property

Public Property Get YAxisPosition() As Long
YAxisPosition = mYAxisPosition
End Property

Public Property Get YAxisWidthCm() As Single
YAxisWidthCm = mYAxisWidthCm
End Property

Public Property Let YAxisWidthCm(ByVal value As Single)
mYAxisWidthCm = value
End Property

'================================================================================
' Methods
'================================================================================

Public Function addChartRegion(ByVal percentheight As Double, _
                    Optional ByVal minimumPercentHeight As Double) As ChartRegion
'
' NB: percentHeight=100 means the region will use whatever space
' is available
'

Dim YAxisRegion As ChartRegion
Dim btn As Button

Set addChartRegion = New ChartRegion

If mRegionsIndex = -1 Then
    addChartRegion.toolbar = Toolbar1
    For Each btn In Toolbar1.Buttons
        btn.Enabled = True
        If btn.key = "showcrosshair" Then btn.value = tbrPressed
    Next
End If

Load ChartRegionPicture(ChartRegionPicture.UBound + 1)
ChartRegionPicture(ChartRegionPicture.UBound).align = vbAlignNone
ChartRegionPicture(ChartRegionPicture.UBound).width = _
    UserControl.ScaleWidth * (mYAxisPosition - chartLeft) / XAxisPicture.ScaleWidth
ChartRegionPicture(ChartRegionPicture.UBound).visible = True

Load YAxisPicture(YAxisPicture.UBound + 1)
YAxisPicture(YAxisPicture.UBound).align = vbAlignNone
YAxisPicture(YAxisPicture.UBound).left = ChartRegionPicture(ChartRegionPicture.UBound).width
YAxisPicture(YAxisPicture.UBound).width = UserControl.ScaleWidth - YAxisPicture(YAxisPicture.UBound).left
YAxisPicture(YAxisPicture.UBound).visible = True

addChartRegion.surface = ChartRegionPicture(ChartRegionPicture.UBound)
addChartRegion.suppressDrawing = mSuppressDrawing
addChartRegion.currentTool = mCurrentTool
addChartRegion.gridColor = mGridColour
addChartRegion.gridTextColor = mGridTextColor
addChartRegion.minimumPercentHeight = minimumPercentHeight
addChartRegion.percentheight = percentheight
addChartRegion.regionBackColor = mBackColour
addChartRegion.regionLeft = mScaleLeft
addChartRegion.regionNumber = mRegionsIndex + 2
addChartRegion.regionBottom = 0
addChartRegion.regionTop = 1
addChartRegion.showCrosshairs = mShowCrosshairs
addChartRegion.showGrid = mShowGrid
addChartRegion.periodsInView mScaleLeft, mYAxisPosition - 1
addChartRegion.autoscale = mAutoscale
addChartRegion.periodLengthMinutes = mPeriodLengthMinutes
addChartRegion.verticalGridUnits = mVerticalGridUnits
addChartRegion.verticalGridSpacing = mVerticalGridSpacing
addChartRegion.sessionStartTime = mSessionStartTime

If mRegionsIndex = UBound(mRegions) Then
    ReDim Preserve mRegions(UBound(mRegions) + 100) As RegionTableEntry
End If

mRegionsIndex = mRegionsIndex + 1
Set mRegions(mRegionsIndex).region = addChartRegion
mRegions(mRegionsIndex).percentheight = percentheight
mRegions(mRegionsIndex).useAvailableSpace = (percentheight = 100#)

If mRegionsIndex <> 0 Then
    Load RegionDividerPicture(mRegionsIndex)
    RegionDividerPicture(mRegionsIndex).visible = True
End If

Set YAxisRegion = New ChartRegion
YAxisRegion.surface = YAxisPicture(YAxisPicture.UBound)
YAxisRegion.regionBottom = 0
YAxisRegion.regionTop = 1
addChartRegion.YAxisRegion = YAxisRegion

If Not sizeRegions Then
    ' can't fit this all in! So remove the added region,
    Set addChartRegion = Nothing
    Set mRegions(mRegionsIndex).region = Nothing
    mRegions(mRegionsIndex).percentheight = 0
    mRegions(mRegionsIndex).actualHeight = 0
    mRegions(mRegionsIndex).useAvailableSpace = False
    mRegionsIndex = mRegionsIndex - 1
    Unload ChartRegionPicture(ChartRegionPicture.UBound)
    Unload RegionDividerPicture(RegionDividerPicture.UBound)
    Unload YAxisPicture(YAxisPicture.UBound)
End If

End Function

Public Function addperiod(ByVal timestamp As Date) As period
Set addperiod = mPeriods.addperiod(timestamp)
End Function

Public Function clearChart()
Dim i As Long

For i = 0 To mRegionsIndex
    mRegions(i).region.clearRegion
    Unload ChartRegionPicture(i + 1)
    Unload YAxisPicture(i + 1)
    If i <> mRegionsIndex Then Unload RegionDividerPicture(i + 1)
Next
If Not mXAxisRegion Is Nothing Then mXAxisRegion.clearRegion
Set mXAxisRegion = Nothing
Set mPeriods = Nothing

initialise
End Function

Public Function refresh()
UserControl.refresh
End Function

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
For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.periodsInView mScaleLeft, mYAxisPosition - 1
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.periodsInView mScaleLeft, mScaleLeft + mScaleWidth
setHorizontalScrollBar
paintAll
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Function calcAvailableHeight() As Long
calcAvailableHeight = XAxisPicture.top - _
                    mRegionsIndex * RegionDividerPicture(0).height - _
                    Toolbar1.height
End Function

Private Sub createXAxisRegion()
Dim aFont As StdFont
Set mXAxisRegion = New ChartRegion
mXAxisRegion.surface = XAxisPicture
mXAxisRegion.periodLengthMinutes = mPeriodLengthMinutes
mXAxisRegion.verticalGridSpacing = mVerticalGridSpacing
mXAxisRegion.verticalGridUnits = mVerticalGridUnits
mXAxisRegion.pointerStyle = PointerNone
mXAxisRegion.regionBackColor = mBackColour
mXAxisRegion.regionBottom = 0
mXAxisRegion.regionTop = 1
mXAxisRegion.sessionStartTime = mSessionStartTime
mXAxisRegion.gridColor = mGridColour
mXAxisRegion.gridTextColor = mGridTextColor
mXAxisRegion.showGrid = False
mXAxisRegion.showGridText = True

Set mXCursorText = mXAxisRegion.addText(LayerNumbers.LayerPointer)
mXCursorText.align = AlignTopLeft
mXCursorText.color = vbRed
mXCursorText.box = True
mXCursorText.boxFillColor = mBackColour
mXCursorText.boxStyle = LineInvisible
Set aFont = New StdFont
aFont.name = "Arial"
aFont.Size = 10
aFont.Underline = False
aFont.Bold = True
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

mXAxisRegion.suppressDrawing = True
mXCursorText.position = mXAxisRegion.newPoint( _
                            periodNumber, _
                            0, _
                            CoordsLogical, _
                            CoordsCounterDistance)

If mPeriodLengthMinutes < 1440 Then
    mXCursorText.text = FormatDateTime(thisPeriod.timestamp, vbShortDate) & _
                        " " & _
                        FormatDateTime(thisPeriod.timestamp, vbShortTime)
Else
    mXCursorText.text = FormatDateTime(thisPeriod.timestamp, vbShortDate)
End If
mXAxisRegion.suppressDrawing = False

End Sub

Private Sub initialise()

mPrevHeight = UserControl.height

ReDim mRegions(100) As RegionTableEntry
mRegionsIndex = -1

Set mPeriods = New periods
mPeriodLengthMinutes = 5
mVerticalGridUnits = TimeHour

mBackColour = vbWhite
mGridColour = &HC0C0C0
mShowGrid = True
mShowCrosshairs = True

mTwipsPerBar = DefaultTwipsPerBar
mScaleHeight = -100
mScaleTop = 100
mYAxisWidthCm = 1.5

mYAxisPosition = 1
resizeX

mAllowHorizontalMouseScrolling = True
mAllowVerticalMouseScrolling = True

HScroll.height = 0

End Sub

Private Sub paintAll()
Dim region As ChartRegion
Dim i As Long

If mSuppressDrawing Then Exit Sub

mNotFirstMouseMove = False

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.paintRegion
Next
If mXAxisRegion Is Nothing Then createXAxisRegion
mXAxisRegion.paintRegion

End Sub

Private Sub resizeX()
Dim newScaleWidth As Single
Dim i As Long
Dim region As ChartRegion

newScaleWidth = CSng(XAxisPicture.width) / CSng(mTwipsPerBar) - 0.5!
mScaleLeft = mYAxisPosition + _
            (mYAxisWidthCm * TwipsPerCm / XAxisPicture.width * newScaleWidth) - _
            newScaleWidth

If newScaleWidth = mScaleWidth Then Exit Sub

mScaleWidth = newScaleWidth

For i = 0 To ChartRegionPicture.UBound
    YAxisPicture(i).left = UserControl.width - YAxisPicture(i).width
    ChartRegionPicture(i).width = YAxisPicture(i).left
Next

For i = 0 To RegionDividerPicture.UBound
    RegionDividerPicture(i).width = UserControl.width
Next

For i = 0 To mRegionsIndex
    Set region = mRegions(i).region
    region.periodsInView mScaleLeft, mYAxisPosition - 1
Next
If Not mXAxisRegion Is Nothing Then
    mXAxisRegion.periodsInView mScaleLeft, mScaleLeft + mScaleWidth
End If

setHorizontalScrollBar
End Sub

Private Sub resizeY()
If UserControl.height = mPrevHeight Then Exit Sub

'debug.print "resizeY"

mPrevHeight = UserControl.height
sizeRegions
End Sub

Private Sub setHorizontalScrollBar()
If mPeriods.currentPeriodNumber + chartWidth - 1 > 32767 Then
    HScroll.Max = 32767
Else
    HScroll.Max = mPeriods.currentPeriodNumber + chartWidth - 1
End If
HScroll.Min = 0
'If HScroll.Max = HScroll.Min Then HScroll.Min = HScroll.Min - 1
HScroll.value = HScroll.Min + Round((CLng(HScroll.Max) - CLng(HScroll.Min)) * lastVisiblePeriod / (mPeriods.currentPeriodNumber + chartWidth - 1))
HScroll.SmallChange = 1
HScroll.LargeChange = chartWidth
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
Dim heightReductionFactor As Double
Dim totalMinimumPercents As Double
Dim nonFixedAvailableSpacePercent As Double
Dim availableSpacePercent As Double
Dim availableHeight As Long     ' the space available for the region picture boxes
                                ' excluding the divider pictures

availableSpacePercent = 100
nonFixedAvailableSpacePercent = 100
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    mRegions(i).percentheight = aRegion.percentheight
    If Not mRegions(i).useAvailableSpace Then
        availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
        nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
    Else
        If aRegion.minimumPercentHeight <> 0 Then
            availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
        End If
        numAvailableSpaceRegions = numAvailableSpaceRegions + 1
    End If
Next

If availableSpacePercent < 0 And mUserResizingRegions Then
    sizeRegions = False
    Exit Function
End If

heightReductionFactor = 1
Do While availableSpacePercent < 0
    availableSpacePercent = 100
    nonFixedAvailableSpacePercent = 100
    heightReductionFactor = heightReductionFactor * 0.66666667
    For i = 0 To mRegionsIndex
        Set aRegion = mRegions(i).region
        If Not mRegions(i).useAvailableSpace Then
            If aRegion.minimumPercentHeight <> 0 Then
                If aRegion.percentheight * heightReductionFactor >= _
                    aRegion.minimumPercentHeight _
                Then
                    mRegions(i).percentheight = aRegion.percentheight * heightReductionFactor
                Else
                    mRegions(i).percentheight = aRegion.minimumPercentHeight
                    totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
                End If
            Else
                mRegions(i).percentheight = aRegion.percentheight * heightReductionFactor
            End If
            availableSpacePercent = availableSpacePercent - mRegions(i).percentheight
            nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - mRegions(i).percentheight
        Else
            If aRegion.minimumPercentHeight <> 0 Then
                availableSpacePercent = availableSpacePercent - aRegion.minimumPercentHeight
                totalMinimumPercents = totalMinimumPercents + aRegion.minimumPercentHeight
            End If
        End If
    Next
    If totalMinimumPercents > 100 Then
        ' can't possibly fit this all in!
        sizeRegions = False
        Exit Function
    End If
Loop

If numAvailableSpaceRegions = 0 Then
    ' we must adjust the percentages on the other regions so they
    ' total 100.
    For i = 0 To mRegionsIndex
        mRegions(i).percentheight = 100 * mRegions(i).percentheight / (100 - nonFixedAvailableSpacePercent)
    Next
End If

' calculate the actual available height to put these regions in
availableHeight = calcAvailableHeight

' first set heights for fixed height regions
For i = 0 To mRegionsIndex
    If Not mRegions(i).useAvailableSpace Then
        mRegions(i).actualHeight = mRegions(i).percentheight * availableHeight / 100
    End If
Next

' now set heights for 'available space' regions with a minimum height
' that needs to be respected
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    If mRegions(i).useAvailableSpace Then
        mRegions(i).actualHeight = 0
        If aRegion.minimumPercentHeight <> 0 Then
            If (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) < aRegion.minimumPercentHeight Then
                mRegions(i).actualHeight = aRegion.minimumPercentHeight * availableHeight / 100
                nonFixedAvailableSpacePercent = nonFixedAvailableSpacePercent - aRegion.minimumPercentHeight
                numAvailableSpaceRegions = numAvailableSpaceRegions - 1
            End If
        End If
    End If
Next

' finally set heights for all other 'available space' regions
For i = 0 To mRegionsIndex
    If mRegions(i).useAvailableSpace And _
        mRegions(i).actualHeight = 0 _
    Then
        mRegions(i).actualHeight = (nonFixedAvailableSpacePercent / numAvailableSpaceRegions) * availableHeight / 100
    End If
Next

' Now actually set the heights and positions for the picture boxes
top = Toolbar1.height
    
For i = 0 To mRegionsIndex
    Set aRegion = mRegions(i).region
    ChartRegionPicture(aRegion.regionNumber).height = mRegions(i).actualHeight
    YAxisPicture(aRegion.regionNumber).height = mRegions(i).actualHeight
    ChartRegionPicture(aRegion.regionNumber).top = top
    YAxisPicture(aRegion.regionNumber).top = top
    top = top + ChartRegionPicture(aRegion.regionNumber).height
    aRegion.resizedY
    If i <> mRegionsIndex Then
        RegionDividerPicture(i + 1).top = top
        top = top + RegionDividerPicture(i + 1).height
    End If
Next

sizeRegions = True
End Function

Private Sub zoom(ByRef rect As TRectangle)

End Sub

