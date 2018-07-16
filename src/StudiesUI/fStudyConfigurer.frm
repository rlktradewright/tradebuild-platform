VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form fStudyConfigurer 
   BorderStyle     =   0  'None
   Caption         =   "Configure a Study"
   ClientHeight    =   12810
   ClientLeft      =   0
   ClientTop       =   -75
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12810
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWButton SetDefaultButton 
      Height          =   615
      Left            =   1320
      TabIndex        =   2
      Top             =   12120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Set as &default"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
   Begin TWControls40.TWButton CancelButton 
      Cancel          =   -1  'True
      Height          =   615
      Left            =   6120
      TabIndex        =   4
      Top             =   12120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Cancel"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
   Begin TWControls40.TWButton AddButton 
      Default         =   -1  'True
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   12120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "&Add to chart"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
   Begin StudiesUI27.StudyConfigurer StudyConfigurer1 
      Height          =   11955
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   21087
   End
   Begin TWControls40.TWButton ApplyDefaultButton 
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   12120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1085
      Caption         =   "Use default"
      DefaultBorderColor=   15793920
      DisabledBackColor=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseOverBackColor=   0
      PushedBackColor =   0
   End
End
Attribute VB_Name = "fStudyConfigurer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

'Event Cancelled()
'Event SetDefault(ByVal studyConfig As studyConfiguration)
'Event AddStudyConfiguration(ByVal studyConfig As studyConfiguration)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "fStudyConfigurer"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                              As Boolean
Private mStudyConfig                            As StudyConfiguration

Private mInitialConfiguration                   As StudyConfiguration

Private mTheme                                  As ITheme

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub Form_Initialize()
StudyConfigurer1.Visible = True
End Sub

Private Sub Form_Load()
Me.Left = 0
Me.Top = 0
Me.Height = Screen.Height
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
Const ProcName As String = "AddButton_Click"
On Error GoTo Err

mCancelled = False
Set mStudyConfig = StudyConfigurer1.StudyConfiguration
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ApplyDefaultButton_Click()
Const ProcName As String = "ApplyDefaultButton_Click"
On Error GoTo Err

StudyConfigurer1.ApplyDefaultConfiguration

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

mCancelled = True
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub SetDefaultButton_Click()
Const ProcName As String = "SetDefaultButton_Click"
On Error GoTo Err

SetDefaultStudyConfiguration StudyConfigurer1.StudyConfiguration

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Cancelled() As Boolean
Const ProcName As String = "Cancelled"
On Error GoTo Err

Cancelled = mCancelled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get StudyConfiguration() As StudyConfiguration
Const ProcName As String = "StudyConfiguration"
On Error GoTo Err

Assert Not mCancelled, "Study configuration was cancelled by the user"

Set StudyConfiguration = mStudyConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChartManager As ChartManager, _
                ByVal pStudyName As String, _
                ByVal pStudyLibraryName As String, _
                ByVal pInitialConfiguration As StudyConfiguration, _
                ByVal pNoParameterModification As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStudyConfig = Nothing
mCancelled = False

Me.Caption = pStudyName

StudyConfigurer1.Initialise _
                pChartManager, _
                pStudyName, _
                pStudyLibraryName, _
                pInitialConfiguration, _
                pNoParameterModification

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



