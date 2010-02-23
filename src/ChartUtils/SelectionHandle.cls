VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectionHandle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "SelectionHandle"

'@================================================================================
' Member variables
'@================================================================================

Private mdp                                 As DataPoint
Attribute mdp.VB_VarHelpID = -1
Private mGraphObj                           As IGraphicObject

Private mStartX                             As Double
Private mStartY                             As Double

Private mAllowXMove                         As Boolean
Private mAllowYMove                         As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get AllowXMove() As Boolean
AllowXMove = mAllowXMove
End Property

Friend Property Get AllowYMove() As Boolean
AllowYMove = mAllowYMove
End Property

Friend Property Get DataPoint() As DataPoint
Set DataPoint = mdp
End Property

Friend Property Let X(ByVal value As Double)
mdp.X = value
End Property

Public Property Get X() As Double
X = mdp.X
End Property

Friend Property Let Y(ByVal value As Double)
mdp.DataValue = value
End Property

Public Property Get Y() As Double
Y = mdp.DataValue
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function HitTest( _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean

Const ProcName As String = "HitTest"
On Error GoTo Err

HitTest = mGraphObj.HitTest(X, Y)

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

Friend Sub Initialise( _
                ByVal dp As DataPoint, _
                ByVal AllowXMove As Boolean, _
                ByVal AllowYMove As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mdp = dp
Set mGraphObj = dp
mAllowXMove = AllowXMove
mAllowYMove = AllowYMove

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub StartMove()
Const ProcName As String = "StartMove"
On Error GoTo Err

mStartX = X
mStartY = Y

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub UndoMove()
Const ProcName As String = "UndoMove"
On Error GoTo Err

X = mStartX
Y = mStartY

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


