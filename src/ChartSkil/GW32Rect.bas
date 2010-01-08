Attribute VB_Name = "GW32Rect"
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

Private Const ModuleName                            As String = "GW32Rect"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function W32RectBottomCentre( _
                ByRef pRect As RECT) As W32Point
W32RectBottomCentre.X = (pRect.Right + pRect.Left) / 2
W32RectBottomCentre.Y = pRect.Bottom
End Function

Public Function W32RectBottomLeft( _
                ByRef pRect As RECT) As W32Point
W32RectBottomLeft.X = pRect.Left
W32RectBottomLeft.Y = pRect.Bottom
End Function

Public Function W32RectBottomRight( _
                ByRef pRect As RECT) As W32Point
W32RectBottomRight.X = pRect.Right
W32RectBottomRight.Y = pRect.Bottom
End Function

Public Function W32RectCentreCentre( _
                ByRef pRect As RECT) As W32Point
W32RectCentreCentre.X = (pRect.Right + pRect.Left) / 2
W32RectCentreCentre.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function W32RectCentreLeft( _
                ByRef pRect As RECT) As W32Point
W32RectCentreLeft.X = pRect.Left
W32RectCentreLeft.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function W32RectCentreRight( _
                ByRef pRect As RECT) As W32Point
W32RectCentreRight.X = pRect.Right
W32RectCentreRight.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function W32RectTopCentre( _
                ByRef pRect As RECT) As W32Point
W32RectTopCentre.X = (pRect.Right + pRect.Left) / 2
W32RectTopCentre.Y = pRect.Top
End Function

Public Function W32RectTopLeft( _
                ByRef pRect As RECT) As W32Point
W32RectTopLeft.X = pRect.Left
W32RectTopLeft.Y = pRect.Top
End Function

Public Function W32RectTopRight( _
                ByRef pRect As RECT) As W32Point
W32RectTopRight.X = pRect.Right
W32RectTopRight.Y = pRect.Top
End Function

'@================================================================================
' Helper Functions
'@================================================================================


