VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#31.0#0"; "TWControls40.ocx"
Begin VB.UserControl StudyPicker 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   4335
   ScaleWidth      =   8655
   Begin TWControls40.TWButton RemoveButton 
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   "<"
      DefaultBorderColor=   15793920
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton AddButton 
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      Caption         =   ">"
      DefaultBorderColor=   15793920
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton ChangeButton 
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Change"
      DefaultBorderColor=   15793920
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton ConfigureButton 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Co&nfigure"
      DefaultBorderColor=   15793920
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView ChartStudiesTree 
      Height          =   2535
      Left            =   3840
      TabIndex        =   2
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      Style           =   2
      SingleSel       =   -1  'True
      Appearance      =   0
   End
   Begin VB.ListBox StudyList 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "StudyPicker.ctx":0000
      Left            =   120
      List            =   "StudyPicker.ctx":0002
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox DescriptionText 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   3480
      Width           =   8415
   End
   Begin VB.Label Label1 
      Caption         =   "Available studies"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Studies in chart"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "StudyPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "StudyPicker"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================


'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mChartManager                As ChartManager
Attribute mChartManager.VB_VarHelpID = -1

Private mAvailableStudies()                     As StudyListEntry

Private mConfigForm                             As fStudyConfigurer
Attribute mConfigForm.VB_VarHelpID = -1

Private mTheme                                  As ITheme

Private mOwner                                  As Variant

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

SendMessage StudyList.hWnd, LB_SETHORZEXTENT, 1000, 0
SendMessage ChartStudiesTree.hWnd, LB_SETHORZEXTENT, 1000, 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = value

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

Dim slName As String
slName = mAvailableStudies(StudyList.ListIndex).StudyLibrary

Dim defaultStudyConfig As StudyConfiguration
Set defaultStudyConfig = mChartManager.GetDefaultStudyConfiguration(mAvailableStudies(StudyList.ListIndex).name, slName)

addStudyToChart defaultStudyConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChangeButton_Click()
Const ProcName As String = "ChangeButton_Click"
On Error GoTo Err

Dim studyConfig As StudyConfiguration
Set studyConfig = mChartManager.GetStudyConfiguration(ChartStudiesTree.SelectedItem.Key)

' NB: the following line displays a modal form, so we can remove the existing
' study and deal with any related studies after it
Dim newStudyConfig As StudyConfiguration
Set newStudyConfig = showConfigForm(studyConfig.name, _
                studyConfig.StudyLibraryName, _
                studyConfig)
If Not newStudyConfig Is Nothing Then
    
    mChartManager.ReplaceStudyConfiguration studyConfig, newStudyConfig
    
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
    DescriptionText = ""
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChartStudiesTree_Click()
Const ProcName As String = "ChartStudiesTree_Click"
On Error GoTo Err

If ChartStudiesTree.SelectedItem Is Nothing Then
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
Else
    ChartStudiesTree.SelectedItem.Expanded = True
    
    Dim studyConfig As StudyConfiguration
    Set studyConfig = mChartManager.GetStudyConfiguration(ChartStudiesTree.SelectedItem.Key)
    
    Dim studyDef As StudyDefinition
    Set studyDef = mChartManager.StudyLibraryManager.GetStudyDefinition( _
                            studyConfig.name, _
                            studyConfig.StudyLibraryName)
    If Not studyDef Is Nothing Then
        StudyList.ListIndex = -1
        AddButton.Enabled = False
        ConfigureButton.Enabled = False
        
        DescriptionText.Text = studyDef.Description
        RemoveButton.Enabled = Not (studyConfig.Study Is mChartManager.BaseStudy)
        ChangeButton.Enabled = True
    End If
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ConfigureButton_Click()
Const ProcName As String = "ConfigureButton_Click"
On Error GoTo Err

Dim studyConfig As StudyConfiguration
Set studyConfig = showConfigForm(mAvailableStudies(StudyList.ListIndex).name, _
                mAvailableStudies(StudyList.ListIndex).StudyLibrary, _
                mChartManager.GetDefaultStudyConfiguration(mAvailableStudies(StudyList.ListIndex).name, _
                                            mAvailableStudies(StudyList.ListIndex).StudyLibrary))
If studyConfig Is Nothing Then Exit Sub
addStudyToChart studyConfig

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RemoveButton_Click()
Const ProcName As String = "RemoveButton_Click"
On Error GoTo Err

Dim studyConfig As StudyConfiguration
Set studyConfig = mChartManager.GetStudyConfiguration(ChartStudiesTree.SelectedItem.Key)
mChartManager.RemoveStudyConfiguration studyConfig
RemoveButton.Enabled = False
ChangeButton.Enabled = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub StudyList_Click()
Const ProcName As String = "StudyList_Click"
On Error GoTo Err

If mChartManager Is Nothing Then Exit Sub

If StudyList.ListIndex <> -1 Then
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
    
    AddButton.Enabled = True
    ConfigureButton.Enabled = True
    
    DescriptionText.Text = mChartManager.StudyLibraryManager.GetStudyDefinition( _
                                    mAvailableStudies(StudyList.ListIndex).name, _
                                    mAvailableStudies(StudyList.ListIndex).StudyLibrary).Description
Else
    AddButton.Enabled = False
    ConfigureButton.Enabled = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mChartManager Event Handlers
'@================================================================================

Private Sub mChartManager_StudyAdded( _
                ByVal pStudy As IStudy)
Const ProcName As String = "mChartManager_StudyAdded"
On Error GoTo Err

Dim parentNode As Node

If Not pStudy.UnderlyingStudy Is Nothing Then
    On Error Resume Next
    Set parentNode = ChartStudiesTree.Nodes.item(pStudy.UnderlyingStudy.Id)
    On Error GoTo Err
End If
If parentNode Is Nothing Then
    ChartStudiesTree.Nodes.Add , _
                                TreeRelationshipConstants.tvwChild, _
                                pStudy.Id, _
                                pStudy.InstanceName
Else
    ChartStudiesTree.Nodes.Add parentNode, _
                                TreeRelationshipConstants.tvwChild, _
                                pStudy.Id, _
                                pStudy.InstanceName
    parentNode.Expanded = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mChartManager_StudyRemoved( _
                ByVal pStudy As IStudy)
Const ProcName As String = "mChartManager_StudyRemoved"
On Error GoTo Err

On Error Resume Next
ChartStudiesTree.Nodes.Remove pStudy.Id

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Parent() As Object
Set Parent = UserControl.Parent
End Property

Public Property Let Theme(ByVal value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

'If mTheme Is value Then Exit Property
Set mTheme = value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

If Not mConfigForm Is Nothing Then mConfigForm.Theme = mTheme

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

Public Sub Initialise( _
                ByVal pChartManager As ChartManager, _
                ByVal pOwner As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mChartManager = pChartManager
gSetVariant mOwner, pOwner

DescriptionText = ""
ChartStudiesTree.Nodes.Clear
StudyList.Clear

If mChartManager Is Nothing Then
ElseIf mChartManager.BaseStudyConfiguration Is Nothing Then
Else
    addEntryToChartStudiesTree mChartManager.BaseStudyConfiguration, Nothing
    
    mAvailableStudies = mChartManager.StudyLibraryManager.GetAvailableStudies
    
    Dim i As Long
    For i = 0 To UBound(mAvailableStudies)
        StudyList.AddItem mAvailableStudies(i).name & "  (" & mAvailableStudies(i).StudyLibrary & ")"
    Next
End If

AddButton.Enabled = False
ConfigureButton.Enabled = False
RemoveButton.Enabled = False
ChangeButton.Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addEntryToChartStudiesTree( _
                ByVal studyConfig As StudyConfiguration, _
                ByVal parentStudyConfig As StudyConfiguration)
Const ProcName As String = "addEntryToChartStudiesTree"
On Error GoTo Err

If studyConfig Is Nothing Then Exit Sub

Dim parentNode As Node
If Not parentStudyConfig Is Nothing Then Set parentNode = ChartStudiesTree.Nodes.item(parentStudyConfig.Study.Id)
If parentNode Is Nothing Then
    ChartStudiesTree.Nodes.Add , _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName
Else
    ChartStudiesTree.Nodes.Add parentNode, _
                                TreeRelationshipConstants.tvwChild, _
                                studyConfig.Study.Id, _
                                studyConfig.Study.InstanceName
    parentNode.Expanded = True
End If

Dim childStudyConfig As StudyConfiguration
For Each childStudyConfig In studyConfig.StudyConfigurations
    addEntryToChartStudiesTree childStudyConfig, studyConfig
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addStudyToChart(ByVal studyConfig As StudyConfiguration)
Const ProcName As String = "addStudyToChart"
On Error GoTo Err

mChartManager.AddStudyConfiguration studyConfig
mChartManager.StartStudy studyConfig.Study

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'/**
'   Returns the required studyConfiguration if the config form is not cancelled by the user
'*/
Private Function showConfigForm( _
                ByVal studyName As String, _
                ByVal slName As String, _
                ByVal initialConfiguration As StudyConfiguration) As StudyConfiguration
Const ProcName As String = "showConfigForm"
On Error GoTo Err

If mConfigForm Is Nothing Then Set mConfigForm = New fStudyConfigurer

Dim noParameterModification  As Boolean
If Not initialConfiguration Is Nothing Then
    If Not mChartManager.BaseStudy Is Nothing Then
        If initialConfiguration.Study Is mChartManager.BaseStudy Then noParameterModification = True
    End If
End If

mConfigForm.Initialise mChartManager, _
                        studyName, _
                        slName, _
                        initialConfiguration, _
                        noParameterModification
If Not mTheme Is Nothing Then mConfigForm.Theme = mTheme
mConfigForm.Show vbModal, mOwner
If Not mConfigForm.Cancelled Then Set showConfigForm = mConfigForm.StudyConfiguration

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





