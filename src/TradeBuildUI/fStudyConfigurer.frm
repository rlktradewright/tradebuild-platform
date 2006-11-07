VERSION 5.00
Begin VB.Form fStudyConfigurer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure a study"
   ClientHeight    =   5745
   ClientLeft      =   990
   ClientTop       =   1215
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   13185
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton SetDefaultButton 
      Caption         =   "Set as &default"
      Height          =   615
      Left            =   12000
      TabIndex        =   3
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   12000
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add to chart"
      Height          =   615
      Left            =   12000
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin TradeBuildUI.StudyConfigurer StudyConfigurer1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   10186
   End
End
Attribute VB_Name = "fStudyConfigurer"
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

Event Cancelled()
Event SetDefault(ByVal studyConfig As studyConfiguration)
Event AddStudyConfiguration(ByVal studyConfig As studyConfiguration)

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

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitCommonControls
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' Control Event Handlers
'================================================================================

Private Sub AddButton_Click()
RaiseEvent AddStudyConfiguration(StudyConfigurer1.studyConfiguration)
Unload Me
End Sub

Private Sub CancelButton_Click()
RaiseEvent Cancelled
Unload Me
End Sub

Private Sub SetDefaultButton_Click()
RaiseEvent SetDefault(StudyConfigurer1.studyConfiguration)
End Sub

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

Friend Sub initialise( _
                ByVal studyDef As TradeBuild.studyDefinition, _
                ByVal serviceProviderName As String, _
                ByRef regionNames() As String, _
                ByVal configuredStudies As StudyConfigurations, _
                ByVal defaultConfiguration As studyConfiguration, _
                ByVal defaultParameters As TradeBuild.parameters)
                
Me.caption = studyDef.name

StudyConfigurer1.initialise _
                studyDef, _
                serviceProviderName, _
                regionNames, _
                configuredStudies, _
                defaultConfiguration, _
                defaultParameters
End Sub

'================================================================================
' Helper Functions
'================================================================================

