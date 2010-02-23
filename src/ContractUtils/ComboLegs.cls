VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboLegs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'
'@================================================================================
' Amendment history
'@================================================================================
'
'
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

Event CollectionChanged( _
                ev As CollectionChangeEvent)

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "ComboLegs"

'@================================================================================
' Member variables
'@================================================================================

Private mComboLegs As Collection

Private mChangeListeners As Collection

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mComboLegs = New Collection
Set mChangeListeners = New Collection
End Sub

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Count() As Long
Count = mComboLegs.Count
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add(ByVal pComboLeg As ComboLeg, _
                Optional ByVal Key As String)
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

If Key <> "" Then
    mComboLegs.Add pComboLeg, Key
Else
    mComboLegs.Add pComboLeg
End If
fireChange CollItemAdded, pComboLeg

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub AddCollectionChangeListener( _
                ByVal value As CollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
Dim failpoint As String
On Error GoTo Err

mChangeListeners.Add value

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Friend Function AddLeg( _
                ByVal ContractSpec As ContractSpecifier, _
                ByVal IsBuyLeg As Boolean, _
                ByVal ratio As Long, _
                ByVal Key As String) As ComboLeg

Const ProcName As String = "AddLeg"
Dim failpoint As String
On Error GoTo Err

Set AddLeg = New ComboLeg
AddLeg.Initialise ContractSpec, IsBuyLeg, ratio
Add AddLeg, Key

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function Item(ByVal index As Variant) As ComboLeg
Attribute Item.VB_UserMemId = 0
Dim theItem As ComboLeg
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set theItem = mComboLegs(index)
Item = theItem

Exit Function

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mComboLegs.[_NewEnum]
End Function

Public Sub Remove(ByVal index As Variant)
Dim lComboLeg As ComboLeg
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

Set lComboLeg = mComboLegs(index)
mComboLegs.Remove index
fireChange CollItemRemoved, lComboLeg

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

Public Sub RemoveCollectionChangeListener(ByVal value As CollectionChangeListener)
Dim i As Long
Const ProcName As String = "RemoveCollectionChangeListener"
Dim failpoint As String
On Error GoTo Err

For i = mChangeListeners.Count To 1 Step -1
    If mChangeListeners.Item(i) Is value Then mChangeListeners.Remove i
Next

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal changeType As CollectionChangeTypes, _
                ByVal affectedObject As ComboLeg)
Dim listener As CollectionChangeListener
Dim i As Long
Dim ev As CollectionChangeEvent
Const ProcName As String = "fireChange"
Dim failpoint As String
On Error GoTo Err

Set ev.Source = Me
ev.changeType = changeType
Set ev.affectedItem = affectedObject
For i = 1 To mChangeListeners.Count
    Set listener = mChangeListeners(i)
    listener.Change ev
Next
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
HandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName, pFailpoint:=failpoint
End Sub



