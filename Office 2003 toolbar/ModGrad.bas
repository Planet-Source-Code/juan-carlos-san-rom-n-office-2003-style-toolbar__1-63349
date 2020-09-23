Attribute VB_Name = "ModGrad"
Option Explicit

' full version of APILine
Public Sub APILineEx(lhdcEx As Long, X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, lcolor As Long)

    'Use the API LineTo for Fast Drawing
    Dim pt As POINT
    Dim hPen As Long, hPenOld As Long
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(lhdcEx, hPen)
    MoveToEx lhdcEx, X1, Y1, pt
    LineTo lhdcEx, X2, Y2
    SelectObject lhdcEx, hPenOld
    DeleteObject hPen
End Sub

Public Function APIRectangle(ByVal hdc As Long, ByVal x As Long, ByVal Y As Long, ByVal w As Long, ByVal h As Long, Optional lcolor As OLE_COLOR = -1) As Long
    
    Dim hPen As Long, hPenOld As Long
    Dim R
    Dim pt As POINT
    hPen = CreatePen(0, 1, lcolor)
    hPenOld = SelectObject(hdc, hPen)
    MoveToEx hdc, x, Y, pt
    LineTo hdc, x + w, Y
    LineTo hdc, x + w, Y + h
    LineTo hdc, x, Y + h
    LineTo hdc, x, Y
    SelectObject hdc, hPenOld
    DeleteObject hPen
End Function

Public Sub DrawVGradientEx(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, ByVal x As Long, ByVal Y As Long, ByVal X2 As Long, ByVal Y2 As Long)
    
    ''Draw a Vertical Gradient in the current HDC
    Dim dR As Single, dG As Single, dB As Single
    Dim sR As Single, sG As Single, sB As Single
    Dim eR As Single, eG As Single, eB As Single
    Dim ni As Long
    sR = (lStartcolor And &HFF)
    sG = (lStartcolor \ &H100) And &HFF
    sB = (lStartcolor And &HFF0000) / &H10000
    eR = (lEndColor And &HFF)
    eG = (lEndColor \ &H100) And &HFF
    eB = (lEndColor And &HFF0000) / &H10000
    dR = (sR - eR) / Y2
    dG = (sG - eG) / Y2
    dB = (sB - eB) / Y2
    For ni = 0 To Y2
        APILineEx lhdcEx, x, Y + ni, X2, Y + ni, RGB(eR + (ni * dR), eG + (ni * dG), eB + (ni * dB))
    Next ni
End Sub

Public Sub DrawGradBorderRect(lhdcEx As Long, lEndColor As Long, lStartcolor As Long, R As RECT, Optional lcolor As OLE_COLOR = -1)
    'draw gradient rectangle with border
    DrawVGradientEx lhdcEx, lEndColor, lStartcolor, R.Left, R.Top, R.Right, R.Bottom
    APIRectangle lhdcEx, R.Left, R.Top, R.Right, R.Bottom, lcolor
End Sub

'Blend two colors
Public Function BlendColors(ByVal lcolor1 As Long, ByVal lcolor2 As Long)
    BlendColors = RGB(((lcolor1 And &HFF) + (lcolor2 And &HFF)) / 2, (((lcolor1 \ &H100) And &HFF) + ((lcolor2 \ &H100) And &HFF)) / 2, (((lcolor1 \ &H10000) And &HFF) + ((lcolor2 \ &H10000) And &HFF)) / 2)
End Function

'System color code to long rgb
Public Function TranslateColor(ByVal lcolor As Long) As Long

    If OleTranslateColor(lcolor, 0, TranslateColor) Then
          TranslateColor = -1
    End If
    
End Function

Public Function DrawTriangle(Pic As Object, ColorFore As Long, BrushColor, poly() As POINT, NumCoords As Long)
    'Dim poly(1 To 3) As COORD, NumCoords As Long,
    Dim hBrush As Long, hRgn As Long
  
    Pic.ForeColor = ColorFore
    ' Polygon function creates unfilled polygon on screen.
    Polygon Pic.hdc, poly(1), NumCoords
    ' Gets stock black brush.
    hBrush = GetStockObject(BrushColor) 'WHITE_BRUSH)
    ' Creates region to fill with color.
    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
    ' If the creation of the region was successful then color.
    If hRgn Then FillRgn Pic.hdc, hRgn, hBrush
    DeleteObject hRgn
'
'    ' Set scalemode to pixels to set up points of triangle.
'    Me.ForeColor = vbBlack
'    ' Assign values to points.
'    poly(1).x = 9:  poly(1).Y = 9
'    poly(2).x = 15: poly(2).Y = 9
'    poly(3).x = 12: poly(3).Y = 12
'    ' Polygon function creates unfilled polygon on screen.
'    Polygon Me.hdc, poly(1), NumCoords
'    ' Gets stock black brush.
'    hBrush = GetStockObject(BLACKBRUSH)
'    ' Creates region to fill with color.
'    hRgn = CreatePolygonRgn(poly(1), NumCoords, ALTERNATE)
'    ' If the creation of the region was successful then color.
'    If hRgn Then FillRgn Me.hdc, hRgn, hBrush

End Function
