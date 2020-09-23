Attribute VB_Name = "ModGrayScale"
Option Explicit

'added by Dennis (dvrdsr) function excerpted vlad memdc class

Public Function PaintIconGrayscale(ByVal Dest_hDC As Long, _
                                    ByVal hIcon As Long, _
                                    Optional ByVal Dest_X As Long, _
                                    Optional ByVal Dest_Y As Long, _
                                    Optional ByVal Dest_Height As Long, _
                                    Optional ByVal Dest_Width As Long) As Boolean
  
  Dim hBMP_Mask  As Long
  Dim hBMP_Image As Long
  Dim hBMP_Prev  As Long
  Dim hIcon_Temp As Long
  Dim hDC_Temp   As Long
  
  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hIcon = 0 Then Exit Function
  
  ' Extract the bitmaps from the icon
  If pvGetIconBitmaps(hIcon, hBMP_Mask, hBMP_Image) = False Then Exit Function
  
  ' Create a memory DC to work with
  hDC_Temp = CreateCompatibleDC(0)
  

  If hDC_Temp = 0 Then GoTo CleanUp
  
  ' Make the image bitmap gradient
  If pvRenderBitmapGrayscale(hDC_Temp, hBMP_Image, 0, 0) = False Then GoTo CleanUp
  
  ' Extract the gradient bitmap out of the DC
  SelectObject hDC_Temp, hBMP_Prev

  
  ' Take the newly gradient bitmap and make a gradient icon from it
  hIcon_Temp = pvCreateIconFromBMP(hBMP_Mask, hBMP_Image)
  If hIcon_Temp = 0 Then GoTo CleanUp
  
  ' Draw the newly created gradient icon onto the specified DC
  If DrawIconEx(Dest_hDC, Dest_X, Dest_Y, hIcon_Temp, Dest_Width, Dest_Height, 0, 0, &H3) <> 0 Then
    PaintIconGrayscale = True
  End If
  
CleanUp:
  
  DestroyIcon hIcon_Temp: hIcon_Temp = 0
  DeleteDC hDC_Temp: hDC_Temp = 0
  DeleteObject hBMP_Mask: hBMP_Mask = 0
  DeleteObject hBMP_Image: hBMP_Image = 0
  
End Function


Private Function pvGetIconBitmaps(ByVal hIcon As Long, _
                               ByRef Return_hBmpMask As Long, _
                               ByRef Return_hBmpImage As Long) As Boolean
  
  Dim TempICONINFO As ICONINFO
  
  If GetIconInfo(hIcon, TempICONINFO) = 0 Then Exit Function
  Return_hBmpMask = TempICONINFO.hbmMask
  Return_hBmpImage = TempICONINFO.hbmColor
  pvGetIconBitmaps = True
  
End Function


Private Function pvRenderBitmapGrayscale(ByVal Dest_hDC As Long, _
                                      ByVal hBitmap As Long, _
                                      Optional ByVal Dest_X As Long, _
                                      Optional ByVal Dest_Y As Long, _
                                      Optional ByVal Srce_X As Long, _
                                      Optional ByVal Srce_Y As Long _
                                      ) As Boolean
  
  Dim TempBITMAP  As BITMAP
  Dim hScreen     As Long
  Dim hDC_Temp    As Long
  Dim hBMP_Prev   As Long
  Dim MyCounterX  As Long
  Dim MyCounterY  As Long
  Dim NewColor    As Long
  Dim hNewPicture As Long
  Dim DeletePic   As Boolean
  
  ' Make sure parameters passed are valid
  If Dest_hDC = 0 Or hBitmap = 0 Then Exit Function
  
  ' Get the handle to the screen DC
  hScreen = GetDC(0)
  If hScreen = 0 Then Exit Function
  
  ' Create a memory DC to work with the picture
  hDC_Temp = CreateCompatibleDC(hScreen)
  If hDC_Temp = 0 Then GoTo CleanUp
  
  ' If the user specifies NOT to alter the original, then make a copy of it to use
    DeletePic = False
    hNewPicture = hBitmap
    
  ' Select the bitmap into the DC
  hBMP_Prev = SelectObject(hDC_Temp, hNewPicture)
  
  ' Get the height / width of the bitmap in pixels
  If GetObjectAPI(hNewPicture, Len(TempBITMAP), TempBITMAP) = 0 Then GoTo CleanUp
  If TempBITMAP.bmHeight <= 0 Or TempBITMAP.bmWidth <= 0 Then GoTo CleanUp
  
  ' Loop through each pixel and conver it to it's grayscale equivelant
  For MyCounterX = 0 To TempBITMAP.bmWidth - 1
    For MyCounterY = 0 To TempBITMAP.bmHeight - 1
      NewColor = GetPixel(hDC_Temp, MyCounterX, MyCounterY)
      If NewColor <> -1 Then
        Select Case NewColor
          ' If the color is already a grey shade, no need to convert it
          Case vbBlack, vbWhite, &H101010, &H202020, &H303030, &H404040, &H505050, &H606060, &H707070, &H808080, &HA0A0A0, &HB0B0B0, &HC0C0C0, &HD0D0D0, &HE0E0E0, &HF0F0F0
            NewColor = NewColor

          Case Else
            NewColor = 0.33 * (NewColor Mod 256) + _
                     0.59 * ((NewColor \ 256) Mod 256) + _
                     0.11 * ((NewColor \ 65536) Mod 256)
            NewColor = RGB(NewColor, NewColor, NewColor)

        End Select
        SetPixel hDC_Temp, MyCounterX, MyCounterY, NewColor
      End If
    Next MyCounterY
  Next MyCounterX
  
  ' Display the picture on the specified hDC
  BitBlt Dest_hDC, Dest_X, Dest_Y, TempBITMAP.bmWidth, TempBITMAP.bmHeight, hDC_Temp, Srce_X, Srce_Y, vbSrcCopy
  
  pvRenderBitmapGrayscale = True
  
CleanUp:
  
  ReleaseDC 0, hScreen: hScreen = 0
  SelectObject hDC_Temp, hBMP_Prev
  DeleteDC hDC_Temp: hDC_Temp = 0
  If DeletePic = True Then
    DeleteObject hNewPicture
    hNewPicture = 0
  End If
  
End Function

Private Function pvCreateIconFromBMP(ByVal hBMP_Mask As Long, _
                                  ByVal hBMP_Image As Long) As Long
  
  Dim TempICONINFO As ICONINFO
  
  If hBMP_Mask = 0 Or hBMP_Image = 0 Then Exit Function
  
  TempICONINFO.fIcon = 1
  TempICONINFO.hbmMask = hBMP_Mask
  TempICONINFO.hbmColor = hBMP_Image
  
  pvCreateIconFromBMP = CreateIconIndirect(TempICONINFO)
  
End Function

'use
'If m_bEnabled Then
'   PaintPicture m_Icon, ix, iy, m_IconSize, m_IconSize
'Else
'   PaintIconGrayscale UserControl.hdc, m_Icon, ix, iy, m_IconSize, m_IconSize
'End If


