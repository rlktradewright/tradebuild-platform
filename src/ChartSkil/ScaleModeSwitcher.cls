VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ScaleModeSwitcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "ScaleModeSwitcher"

'@================================================================================
' Member variables
'@================================================================================

Private mSurface As PictureBox
Private mScaleLeft As Double
Private mScaleWidth As Double
Private mScaleTop As Double
Private mScaleHeight As Double

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pSurface As PictureBox)
Set mSurface = pSurface
End Sub

Friend Sub SwitchToPixelMode()
Const ProcName As String = "SwitchToPixelMode"
Dim failpoint As String
On Error GoTo Err

mScaleLeft = mSurface.ScaleLeft
mScaleWidth = mSurface.ScaleWidth
mScaleTop = mSurface.ScaleTop
mScaleHeight = mSurface.ScaleHeight

mSurface.ScaleMode = vbPixels
Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

Friend Sub SwitchToUserMode()
Const ProcName As String = "SwitchToUserMode"
Dim failpoint As String
On Error GoTo Err

mSurface.Scale (mScaleLeft, mScaleTop)-(mScaleLeft + mScaleWidth, mScaleTop + mScaleHeight)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


