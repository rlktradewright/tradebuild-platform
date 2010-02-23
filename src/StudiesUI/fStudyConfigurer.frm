VERSION 5.00
Begin VB.Form fStudyConfigurer 
   Caption         =   "Configure a Study"
   ClientHeight    =   5805
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   ScaleHeight     =   5805
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin StudiesUI26.StudyConfigurer StudyConfigurer1 
      Height          =   5655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   12255
      _ExtentX        =   21616
      _ExtentY        =   9975
   End
   Begin VB.CommandButton AddButton 
      Caption         =   "&Add to chart"
      Height          =   615
      Left            =   12360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   615
      Left            =   12360
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton SetDefaultButton 
      Caption         =   "Set as &default"
      Height          =   615
      Left            =   12360
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
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

'@================================================================================
' Form Event Handlers
'@================================================================================

Private Sub Form_Initialize()
InitCommonControls
StudyConfigurer1.Visible = True
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

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
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

mCancelled = True
Me.Hide

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
End Sub

Private Sub SetDefaultButton_Click()
Const ProcName As String = "SetDefaultButton_Click"
On Error GoTo Err

SetDefaultStudyConfiguration StudyConfigurer1.StudyConfiguration

Exit Sub

Err:
UnhandledErrorHandler.Notify ProcName, ModuleName, ProjectName
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
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

Public Property Get StudyConfiguration() As StudyConfiguration
Const ProcName As String = "StudyConfiguration"
On Error GoTo Err

If mCancelled Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Study configuration was cancelled by the user"
End If

Set StudyConfiguration = mStudyConfig

Exit Property

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pChart As ChartController, _
                ByVal studyDef As StudyDefinition, _
                ByVal StudyLibraryName As String, _
                ByRef regionNames() As String, _
                ByRef baseStudyConfig As StudyConfiguration, _
                ByVal defaultConfiguration As StudyConfiguration, _
                ByVal defaultParameters As Parameters, _
                ByVal noParameterModification As Boolean)
                
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mStudyConfig = Nothing
mCancelled = False

Me.Caption = studyDef.name

StudyConfigurer1.Initialise _
                pChart, _
                studyDef, _
                StudyLibraryName, _
                regionNames, _
                baseStudyConfig, _
                defaultConfiguration, _
                defaultParameters, _
                noParameterModification

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



