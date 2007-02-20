VERSION 5.00
Begin VB.Form fStudyConfigurer 
   Caption         =   "Configure a Study"
   ClientHeight    =   5775
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13170
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   13170
   StartUpPosition =   3  'Windows Default
   Begin StudiesUI25.StudyConfigurer StudyConfigurer1 
      Height          =   5535
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   9763
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add to chart"
      Height          =   615
      Left            =   12000
      TabIndex        =   0
      Top             =   120
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
   Begin VB.CommandButton SetDefaultButton 
      Caption         =   "Set as &default"
      Height          =   615
      Left            =   12000
      TabIndex        =   1
      Top             =   1560
      Width           =   1095
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
StudyConfigurer1.Visible = True
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
                ByVal studyDef As StudyDefinition, _
                ByVal StudyLibraryName As String, _
                ByRef regionNames() As String, _
                ByVal configuredStudies As StudyConfigurations, _
                ByVal defaultConfiguration As studyConfiguration, _
                ByVal defaultParameters As Parameters2.Parameters)
                
Me.Caption = studyDef.name

StudyConfigurer1.initialise _
                studyDef, _
                StudyLibraryName, _
                regionNames, _
                configuredStudies, _
                defaultConfiguration, _
                defaultParameters
End Sub

'================================================================================
' Helper Functions
'================================================================================



