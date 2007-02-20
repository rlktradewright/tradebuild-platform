Attribute VB_Name = "Globals"
Option Explicit

Public Declare Function SendMessageByNum Lib "user32" _
    Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub InitCommonControls Lib "comctl32" ()

Public Const LB_SETHORZEXTENT = &H194

