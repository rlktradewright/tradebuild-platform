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

Public Function W32PointAdd( _
                ByRef pPoint1 As W32Point, _
                ByRef pPoint2 As W32Point) As W32Point
W32PointAdd.X = pPoint1.X + pPoint2.X
W32PointAdd.Y = pPoint1.Y + pPoint2.Y
End Function

Public Function W32PointSubtract( _
                ByRef pPoint1 As W32Point, _
                ByRef pPoint2 As W32Point) As W32Point
W32PointSubtract.X = pPoint1.X - pPoint2.X
W32PointSubtract.Y = pPoint1.Y - pPoint2.Y
End Function

Public Function W32PointToString( _
                ByRef pPoint As W32Point) As String
W32PointToString = "X=" & pPoint.X & "; Y=" & pPoint.Y
End Function

Public Sub W32RectAdjustForRotationAboutPoint( _
                ByRef pRect As RECT, _
                ByVal pAngle As Double, _
                ByRef pPoint As W32Point)

OffsetRect pRect, -pPoint.X, -pPoint.Y
W32RectRotate pRect, pAngle
OffsetRect pRect, pPoint.X, pPoint.Y
End Sub

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

Public Sub W32RectRotate( _
                ByRef pRect As RECT, _
                ByVal pAngle As Double)
Dim transform As XForm
Dim p1 As W32Point
Dim p2 As W32Point
Dim p3 As W32Point
Dim p4 As W32Point


transform.eM11 = Cos(-pAngle)
transform.eM12 = Sin(-pAngle)
transform.eM21 = -Sin(-pAngle)
transform.eM22 = Cos(-pAngle)
transform.eDx = 0
transform.eDy = 0

p1 = transformPoint(W32RectBottomLeft(pRect), transform)
p2 = transformPoint(W32RectBottomRight(pRect), transform)
p3 = transformPoint(W32RectTopLeft(pRect), transform)
p4 = transformPoint(W32RectTopRight(pRect), transform)

pRect.Bottom = max4(p1.Y, p2.Y, p3.Y, p4.Y)
pRect.Left = min4(p1.X, p2.X, p3.X, p4.X)
pRect.Top = min4(p1.Y, p2.Y, p3.Y, p4.Y)
pRect.Right = max4(p1.X, p2.X, p3.X, p4.X)
End Sub

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

Public Function W32RectToString( _
                ByRef pRect As RECT) As String
W32RectToString = "Bottom=" & pRect.Bottom & "; Left=" & pRect.Left & "; Top=" & pRect.Top & "; Right=" & pRect.Right
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function max4( _
                ByVal v1 As Long, _
                ByVal v2 As Long, _
                ByVal v3 As Long, _
                ByVal v4 As Long) As Long
max4 = v1
If v2 > max4 Then max4 = v2
If v3 > max4 Then max4 = v3
If v4 > max4 Then max4 = v4
End Function

Private Function min4( _
                ByVal v1 As Long, _
                ByVal v2 As Long, _
                ByVal v3 As Long, _
                ByVal v4 As Long) As Long
min4 = v1
If v2 < min4 Then min4 = v2
If v3 < min4 Then min4 = v3
If v4 < min4 Then min4 = v4
End Function

Private Function transformPoint( _
                ByRef pPoint As W32Point, _
                ByRef pTransform As XForm) As W32Point

transformPoint.X = pPoint.X * pTransform.eM11 - pPoint.Y * pTransform.eM21 + pTransform.eDx
transformPoint.Y = pPoint.X * pTransform.eM12 + pPoint.Y * pTransform.eM22 + pTransform.eDy
End Function

