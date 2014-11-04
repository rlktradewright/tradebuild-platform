VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.UserControl StudyPicker 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   4335
   ScaleWidth      =   8655
   Begin MSComctlLib.TreeView ChartStudiesTree 
      Height          =   2535
      Left            =   3840
      TabIndex        =   9
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   4471
      _Version        =   393217
      HideSelection   =   0   'False
      LabelEdit       =   1
      SingleSel       =   -1  'True
      Appearance      =   0
   End
   Begin VB.ListBox StudyList 
      Height          =   2400
      ItemData        =   "StudyPicker.ctx":0000
      Left            =   120
      List            =   "StudyPicker.ctx":0002
      TabIndex        =   5
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox DescriptionText 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3480
      Width           =   8415
   End
   Begin VB.CommandButton AddButton 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      ToolTipText     =   "Add study to chart"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton ConfigureButton 
      Caption         =   "Co&nfigure"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      ToolTipText     =   "Configure selected study"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton RemoveButton 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Remove study from chart"
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton ChangeButton 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   0
      ToolTipText     =   "Change selected study's configuration"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Available studies"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Studies in chart"
      Height          =   255
      Left            =   3960
      TabIndex        =   6
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "StudyPicker"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================


'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mChartManager As ChartManager
Attribute mChartManager.VB_VarHelpID = -1

Private mAvailableStudies() As StudyListEntry

Private mConfigForm As fStudyConfigurer
Attribute mConfigForm.VB_VarHelpID = -1

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
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
Const ProcName As String = "AddButton_Click"
On Error GoTo Err

Dim slName As String
slName = mAvailableStudies(StudyList.ListIndex).StudyLibrary

Dim defaultStudyConfig As StudyConfiguration
Set defaultStudyConfig = mChartManager.GetDefaultStudyConfiguration(mAvailableStudies(StudyList.ListIndex).name, slName)

If Not defaultStudyConfig Is Nothing Then
    addStudyToChart defaultStudyConfig
Else
    Dim studyConfig As StudyConfiguration
    Set studyConfig = showConfigForm(mAvailableStudies(StudyList.ListIndex).name, _
                mAvailableStudies(StudyList.ListIndex).StudyLibrary, _
                defaultStudyConfig)
    If studyConfig Is Nothing Then Exit Sub
    addStudyToChart studyConfig
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ChangeButton_Click()
Const ProcName As String = "ChangeButton_Click"
On Error GoTo Err

Dim studyConfig As StudyConfiguration
Set studyConfig = mChartManager.GetStudyConfig(ChartStudiesTree.SelectedItem.Key)

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
    Set studyConfig = mChartManager.GetStudyConfig(ChartStudiesTree.SelectedItem.Key)
    
    Dim studyDef As StudyDefinition
    Set studyDef = mChartManager.StudyLibraryManager.GetStudyDefinition( _
                            studyConfig.name, _
                            studyConfig.StudyLibraryName)
    If Not studyDef Is Nothing Then
        StudyList.ListIndex = -1
        AddButton.Enabled = False
        ConfigureButton.Enabled = False
        
        DescriptionText.text = studyDef.Description
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
Set studyConfig = mChartManager.GetStudyConfig(ChartStudiesTree.SelectedItem.Key)
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
    
    DescriptionText.text = mChartManager.StudyLibraryManager.GetStudyDefinition( _
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
                ByVal studyConfig As ChartUtils27.StudyConfiguration)
Const ProcName As String = "mChartManager_StudyAdded"
On Error GoTo Err

Dim parentNode As Node

If Not studyConfig.UnderlyingStudy Is Nothing Then
    On Error Resume Next
    Set parentNode = ChartStudiesTree.Nodes.item(studyConfig.UnderlyingStudy.Id)
    On Error GoTo Err
End If
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

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub mChartManager_StudyRemoved( _
                ByVal studyConfig As ChartUtils27.StudyConfiguration)
Const ProcName As String = "mChartManager_StudyRemoved"
On Error GoTo Err

On Error Resume Next
ChartStudiesTree.Nodes.Remove studyConfig.Study.Id

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pChartManager As ChartManager)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mChartManager = pChartManager

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
                ByVal defaultConfiguration As StudyConfiguration) As StudyConfiguration
Const ProcName As String = "showConfigForm"
On Error GoTo Err

If mConfigForm Is Nothing Then Set mConfigForm = New fStudyConfigurer

Dim noParameterModification  As Boolean
If Not defaultConfiguration Is Nothing Then
    If Not mChartManager.BaseStudy Is Nothing Then
        If defaultConfiguration.Study Is mChartManager.BaseStudy Then noParameterModification = True
    End If
End If

mConfigForm.Initialise mChartManager.Chart, _
                        mChartManager.StudyLibraryManager.GetStudyDefinition(studyName, slName), _
                        slName, _
                        mChartManager.regionNames, _
                        mChartManager.BaseStudyConfiguration, _
                        defaultConfiguration, _
                        mChartManager.StudyLibraryManager.GetStudyDefaultParameters(studyName, slName), _
                        noParameterModification
mConfigForm.Show vbModal, UserControl.Parent
If Not mConfigForm.Cancelled Then Set showConfigForm = mConfigForm.StudyConfiguration

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function





