VERSION 5.00
Begin VB.UserControl StudyPicker 
   ClientHeight    =   4335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   4335
   ScaleWidth      =   8655
   Begin VB.ListBox StudyList 
      Height          =   2595
      ItemData        =   "StudyPicker.ctx":0000
      Left            =   120
      List            =   "StudyPicker.ctx":0002
      TabIndex        =   6
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox DescriptionText 
      Height          =   735
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3480
      Width           =   8415
   End
   Begin VB.CommandButton AddButton 
      Caption         =   ">"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Add study to chart"
      Top             =   1080
      Width           =   375
   End
   Begin VB.CommandButton ConfigureButton 
      Caption         =   "Co&nfigure"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      ToolTipText     =   "Configure selected study"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.ListBox ChartStudiesList 
      Height          =   2595
      ItemData        =   "StudyPicker.ctx":0004
      Left            =   3840
      List            =   "StudyPicker.ctx":0006
      TabIndex        =   2
      Top             =   360
      Width           =   4695
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

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================


'@================================================================================
' Member variables
'@================================================================================

Private mChartManager As ChartManager
Private mChartController As chartController

Private mAvailableStudies() As StudyListEntry

Private WithEvents mStudyConfigurations As StudyConfigurations
Attribute mStudyConfigurations.VB_VarHelpID = -1

Private WithEvents mConfigForm As fStudyConfigurer
Attribute mConfigForm.VB_VarHelpID = -1

''
'   Set in the Study Configuration Form's AddStudyConfiguration event
'@/
Private mNewStudyConfiguration As studyConfiguration

'@================================================================================
' UserControl Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
InitCommonControls
SendMessage StudyList.hWnd, LB_SETHORZEXTENT, 1000, 0
SendMessage ChartStudiesList.hWnd, LB_SETHORZEXTENT, 1000, 0
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub AddButton_Click()
Dim spName As String
Dim defaultStudyConfig As studyConfiguration

spName = mAvailableStudies(StudyList.ListIndex).StudyLibrary
Set defaultStudyConfig = loadDefaultStudyConfiguration(mAvailableStudies(StudyList.ListIndex).name, spName)

If Not defaultStudyConfig Is Nothing Then
    addStudyToChart defaultStudyConfig
Else
    showConfigForm mAvailableStudies(StudyList.ListIndex).name, _
                mAvailableStudies(StudyList.ListIndex).StudyLibrary, _
                defaultStudyConfig
End If

mChartController.suppressDrawing = False
End Sub

Private Sub ChangeButton_Click()
Dim studyConfig As studyConfiguration
Dim newStudyConfig As studyConfiguration

Set studyConfig = mStudyConfigurations.item(ChartStudiesList.List(ChartStudiesList.ListIndex))

' NB: the following line displays a modal form, so we can remove the existing
' study and deal with any related studies after it
Set newStudyConfig = showConfigForm(studyConfig.name, _
                studyConfig.StudyLibraryName, _
                studyConfig)
If Not newStudyConfig Is Nothing Then
    
    mChartManager.removeStudy studyConfig
    
    ' now amend any studies that are based on the changed study
    reconfigureDependingStudies studyConfig, newStudyConfig
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
    ChartStudiesList.ListIndex = -1
    DescriptionText = ""
End If
mChartController.suppressDrawing = False

End Sub

Private Sub ChartStudiesList_Click()
Dim studyDef As StudyDefinition
Dim studyConfig As studyConfiguration

If ChartStudiesList.ListIndex < 1 Then
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
Else
    Set studyConfig = mStudyConfigurations.item(ChartStudiesList.List(ChartStudiesList.ListIndex))
    Set studyDef = GetStudyDefinition( _
                            studyConfig.name, _
                            studyConfig.StudyLibraryName)
    If Not studyDef Is Nothing Then
        StudyList.ListIndex = -1
        AddButton.Enabled = False
        ConfigureButton.Enabled = False
        
        DescriptionText.text = studyDef.Description
        RemoveButton.Enabled = True
        ChangeButton.Enabled = True
    End If
End If
End Sub

Private Sub ConfigureButton_Click()
showConfigForm mAvailableStudies(StudyList.ListIndex).name, _
                mAvailableStudies(StudyList.ListIndex).StudyLibrary, _
                Nothing
mChartController.suppressDrawing = False
End Sub

Private Sub RemoveButton_Click()
Dim studyConfig As studyConfiguration
Set studyConfig = mStudyConfigurations.item(ChartStudiesList.List(ChartStudiesList.ListIndex))
mChartManager.removeStudy studyConfig
removeDependingStudies studyConfig
RemoveButton.Enabled = False
ChangeButton.Enabled = False
ChartStudiesList.ListIndex = -1
End Sub

Private Sub StudyList_Click()
Dim studyDef As StudyDefinition
Dim spName As String

If mChartManager Is Nothing Then Exit Sub

If StudyList.ListIndex <> -1 Then
    ChartStudiesList.ListIndex = -1
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
    
    AddButton.Enabled = True
    ConfigureButton.Enabled = True
    spName = mAvailableStudies(StudyList.ListIndex).StudyLibrary
    Set studyDef = GetStudyDefinition( _
                            mAvailableStudies(StudyList.ListIndex).name, _
                            spName)
    DescriptionText.text = studyDef.Description
Else
    AddButton.Enabled = False
    ConfigureButton.Enabled = False
End If
End Sub

'@================================================================================
' mConfigForm Event Handlers
'@================================================================================

Private Sub mConfigForm_Cancelled()
End Sub

Private Sub mConfigForm_SetDefault( _
                ByVal studyConfig As studyConfiguration)
updateDefaultStudyConfiguration studyConfig
End Sub

Private Sub mConfigForm_AddStudyConfiguration( _
                ByVal studyConfig As studyConfiguration)
Set mNewStudyConfiguration = studyConfig
If studyConfig.StudyValueConfigurations.Count = 0 Then Exit Sub

addStudyToChart studyConfig
End Sub

'@================================================================================
' mStudyConfigurations Event Handlers
'@================================================================================

Private Sub mStudyConfigurations_ItemAdded( _
                ByVal studyConfig As studyConfiguration)
ChartStudiesList.AddItem studyConfig.instanceFullyQualifiedName
End Sub

Private Sub mStudyConfigurations_ItemRemoved( _
                ByVal studyConfig As studyConfiguration)
Dim i As Long
For i = 0 To ChartStudiesList.ListCount - 1
    If ChartStudiesList.List(i) = studyConfig.instanceFullyQualifiedName Then
        ChartStudiesList.RemoveItem i
        Exit For
    End If
Next
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub initialise( _
                ByVal pChartManager As ChartManager)
Dim studyConfig As studyConfiguration
Dim i As Long
Dim itemText As String
Dim lLogger As Logger

Set lLogger = GetLogger("diag.tradebuild.studiesui")
lLogger.Log LogLevelMediumDetail, "initialise"

Set mChartManager = pChartManager

DescriptionText = ""
ChartStudiesList.clear
If Not mChartManager Is Nothing Then
    Set mChartController = mChartManager.chartController
    Set mStudyConfigurations = mChartManager.StudyConfigurations
    lLogger.Log LogLevelMediumDetail, "Study configurations: " & mStudyConfigurations.Count
    For Each studyConfig In mStudyConfigurations
        ChartStudiesList.AddItem studyConfig.instanceFullyQualifiedName
    Next
Else
    lLogger.Log LogLevelMediumDetail, "Chart manager is Nothing"
End If

StudyList.clear
mAvailableStudies = AvailableStudies

For i = 0 To UBound(mAvailableStudies)
    itemText = mAvailableStudies(i).name & "  (" & mAvailableStudies(i).StudyLibrary & ")"
    StudyList.AddItem itemText
Next

AddButton.Enabled = False
ConfigureButton.Enabled = False
RemoveButton.Enabled = False
ChangeButton.Enabled = False

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addStudyToChart(ByVal studyConfig As studyConfiguration)
On Error Resume Next
mChartController.suppressDrawing = True
mChartManager.addStudy studyConfig
mChartManager.startStudy studyConfig.study
If err.Number <> 0 Then initialise Nothing
On Error GoTo 0
' don't unsuppress drawing here because there may be a removeStudy to do first
End Sub

'/**
'   Reconfigures any studies that are dependant on the
'   oldStudyConfig to use the NewStudyConfig.
'
' @param oldStudyConfig     a <code>StudyConfiguration</code> object
'                           whose id is used to find other
'                           <code>StudyConfiguration</code> object that
'                           depend on it
'
' @param newStudyConfig     a <code>StudyConfiguration</code> object that
'                           the dependant <code>StudyConfiguration</code>s
'                           must be reconfigured to depend on
'
'*/
Private Sub reconfigureDependingStudies( _
                ByVal oldStudyConfig As studyConfiguration, _
                ByVal newStudyConfig As studyConfiguration)
Dim sc As studyConfiguration
Dim newSc As studyConfiguration

For Each sc In mStudyConfigurations
    If sc.underlyingStudy Is oldStudyConfig.study Then
        Set newSc = sc.Clone
        newSc.underlyingStudy = newStudyConfig.study
        If sc.chartRegionName = oldStudyConfig.chartRegionName Then
            newSc.chartRegionName = newStudyConfig.chartRegionName
        End If
        mChartManager.addStudy newSc
        mChartManager.removeStudy sc
        reconfigureDependingStudies sc, newSc
    End If
Next

End Sub

'/**
'   Removes any studies that are dependant on the
'   specified <code>StudyConfiguration</code>
'
' @param studyConfig    the <code>StudyConfiguration</code> object
'                       whose depending studies are to be removed
'
'*/
Private Sub removeDependingStudies( _
                ByVal studyConfig As studyConfiguration)
Dim sc As studyConfiguration
                
For Each sc In mStudyConfigurations
    If sc.underlyingStudy Is studyConfig.study Then
        mChartManager.removeStudy sc
        removeDependingStudies sc
    End If
Next

End Sub

'/**
'   Returns the required studyConfiguration if the config form is not cancelled by the user
'*/
Private Function showConfigForm( _
                ByVal studyName As String, _
                ByVal spName As String, _
                ByVal defaultConfiguration As studyConfiguration) As studyConfiguration

Set mConfigForm = New fStudyConfigurer

'mConfigForm.Show vbModeless, Me

mConfigForm.initialise mChartManager.chartController, _
                        GetStudyDefinition(studyName, spName), _
                        spName, _
                        mChartManager.regionNames, _
                        mStudyConfigurations, _
                        defaultConfiguration, _
                        GetStudyDefaultParameters(studyName, spName)
mConfigForm.Show vbModal, Me
Set showConfigForm = mNewStudyConfiguration
Set mNewStudyConfiguration = Nothing
End Function





