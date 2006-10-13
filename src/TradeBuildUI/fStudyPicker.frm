VERSION 5.00
Begin VB.Form fStudyPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select a study"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ChangeButton 
      Caption         =   "Change"
      Enabled         =   0   'False
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      ToolTipText     =   "Change selected study's configuration"
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton RemoveButton 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      ToolTipText     =   "Remove study from chart"
      Top             =   1560
      Width           =   375
   End
   Begin VB.ListBox ChartStudiesList 
      Height          =   2595
      ItemData        =   "fStudyPicker.frx":0000
      Left            =   3840
      List            =   "fStudyPicker.frx":0002
      TabIndex        =   7
      Top             =   360
      Width           =   4695
   End
   Begin VB.CommandButton CloseButton 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   7440
      TabIndex        =   3
      ToolTipText     =   "Close this dialog"
      Top             =   3840
      Width           =   1095
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
   Begin VB.CommandButton AddButton 
      Caption         =   ">"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      ToolTipText     =   "Add study to chart"
      Top             =   1080
      Width           =   375
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
      Width           =   7095
   End
   Begin VB.ListBox StudyList 
      Height          =   2595
      ItemData        =   "fStudyPicker.frx":0004
      Left            =   120
      List            =   "fStudyPicker.frx":0006
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "Studies in chart"
      Height          =   255
      Left            =   3960
      TabIndex        =   8
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Available studies"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "fStudyPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================


'================================================================================
' Member variables
'================================================================================

Private mTicker As TradeBuild.ticker
Private mChart As TradeBuildChart

Private mStudies() As TradeBuild.StudyListEntry

Private WithEvents mStudyConfigurations As studyConfigurations
Attribute mStudyConfigurations.VB_VarHelpID = -1

'Private mStudyConfiguration As StudyConfiguration

Private WithEvents mConfigForm As fStudyConfigurer
Attribute mConfigForm.VB_VarHelpID = -1

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
SendMessageByNum StudyList.hwnd, LB_SETHORZEXTENT, 1000, 0
SendMessageByNum ChartStudiesList.hwnd, LB_SETHORZEXTENT, 1000, 0
End Sub

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddButton_Click()
Dim spName As String
Dim defaultStudyConfig As StudyConfiguration

spName = mStudies(StudyList.ListIndex).serviceProvider
Set defaultStudyConfig = loadDefaultStudyConfiguration(mStudies(StudyList.ListIndex).name, spName)

If Not defaultStudyConfig Is Nothing Then
    addStudyToChart defaultStudyConfig
Else
    showConfigForm
End If

End Sub

Private Sub ChangeButton_Click()
notImplemented
End Sub

Private Sub ChartStudiesList_Click()
Dim studyDef As studyDefinition
Dim studyConfig As StudyConfiguration

If ChartStudiesList.ListIndex <> -1 Then
    Set studyConfig = mStudyConfigurations.item(ChartStudiesList.List(ChartStudiesList.ListIndex))
    Set studyDef = mTicker.studyDefinition( _
                            studyConfig.name, _
                            studyConfig.serviceProviderName)
    If Not studyDef Is Nothing Then
        DescriptionText.text = studyDef.Description
        RemoveButton.Enabled = True
        ChangeButton.Enabled = True
        AddButton.Enabled = False
        ConfigureButton.Enabled = False
    End If
Else
    AddButton.Enabled = False
    ConfigureButton.Enabled = False
    DescriptionText.text = ""
End If
End Sub

Private Sub CloseButton_Click()
Unload Me
End Sub

Private Sub ConfigureButton_Click()
showConfigForm
End Sub

Private Sub RemoveButton_Click()
notImplemented
End Sub

Private Sub StudyList_Click()
Dim studyDef As studyDefinition
Dim spName As String

If StudyList.ListIndex <> -1 Then
    AddButton.Enabled = True
    ConfigureButton.Enabled = True
    RemoveButton.Enabled = False
    ChangeButton.Enabled = False
    spName = mStudies(StudyList.ListIndex).serviceProvider
    Set studyDef = mTicker.studyDefinition( _
                            mStudies(StudyList.ListIndex).name, _
                            spName)
    DescriptionText.text = studyDef.Description
Else
    AddButton.Enabled = False
    ConfigureButton.Enabled = False
    DescriptionText.text = ""
End If
End Sub

'================================================================================
' mConfigForm Event Handlers
'================================================================================

Private Sub mConfigForm_SetDefault( _
                ByVal studyConfig As StudyConfiguration)
updateDefaultStudyConfiguration studyConfig
End Sub

Private Sub mConfigForm_AddStudyConfiguration( _
                ByVal studyConfig As StudyConfiguration)
If studyConfig.studyValueConfigurations.count = 0 Then Exit Sub

addStudyToChart studyConfig

End Sub

'================================================================================
' mStudyConfigurations Event Handlers
'================================================================================

Private Sub mStudyConfigurations_ItemAdded( _
                ByVal studyConfig As StudyConfiguration)
ChartStudiesList.AddItem studyConfig.instanceFullyQualifiedName
End Sub

Private Sub mStudyConfigurations_ItemRemoved( _
                ByVal studyConfig As StudyConfiguration)
Dim i As Long
For i = 0 To ChartStudiesList.ListCount - 1
    If ChartStudiesList.List(i) = studyConfig.instanceName Then
        ChartStudiesList.RemoveItem i
        Exit For
    End If
Next
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal pChart As TradeBuildChart, _
                ByVal pTicker As TradeBuild.ticker)
Dim studyConfig As StudyConfiguration
Dim i As Long
Dim itemText As String

Set mChart = pChart
Set mTicker = pTicker

DescriptionText = ""
ChartStudiesList.clear
If Not mChart Is Nothing Then
    Set mStudyConfigurations = mChart.studyConfigurations
    For Each studyConfig In mStudyConfigurations
        ChartStudiesList.AddItem studyConfig.instanceFullyQualifiedName
    Next
End If

StudyList.clear
If Not mTicker Is Nothing Then
    mStudies = mTicker.availableStudies
    
    For i = 0 To UBound(mStudies)
        itemText = mStudies(i).name & "  (" & mStudies(i).serviceProvider & ")"
        StudyList.AddItem itemText
    Next
End If

AddButton.Enabled = False
ConfigureButton.Enabled = False
RemoveButton.Enabled = False
ChangeButton.Enabled = False

If mTicker Is Nothing Or mChart Is Nothing Then
    Me.caption = "(No chart selected)"
Else
    Me.caption = "Select a study for " & mTicker.Contract.specifier.localSymbol & _
                " (" & mTicker.Contract.specifier.exchange & ") " & _
                mChart.timeframeCaption
End If
End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addStudyToChart(ByVal studyConfig As StudyConfiguration)
On Error Resume Next
mChart.addStudy studyConfig
If err.Number <> 0 Then initialise Nothing, Nothing
On Error GoTo 0
End Sub

Private Sub showConfigForm()
Dim spName As String

Set mConfigForm = New fStudyConfigurer

spName = mStudies(StudyList.ListIndex).serviceProvider

mConfigForm.initialise mTicker, _
                        mTicker.studyDefinition( _
                            mStudies(StudyList.ListIndex).name, _
                            spName), _
                        spName, _
                        mChart.regionNames, _
                        mStudyConfigurations, _
                        loadDefaultStudyConfiguration(mStudies(StudyList.ListIndex).name, spName)
mConfigForm.Show vbModal, Me
End Sub



