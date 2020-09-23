VERSION 5.00
Begin VB.UserControl JCToolbar 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6900
   ControlContainer=   -1  'True
   ScaleHeight     =   28
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   460
   ToolboxBitmap   =   "JCToolbar.ctx":0000
   Begin VB.PictureBox PicRight 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   6570
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   15
      TabIndex        =   1
      Top             =   0
      Width           =   230
   End
   Begin VB.PictureBox PicLeft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   0
      MousePointer    =   15  'Size All
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   2
      Top             =   0
      Width           =   135
   End
   Begin VB.Timer tmrRight 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1770
      Top             =   540
   End
   Begin VB.PictureBox PicGrad 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   210
      ScaleHeight     =   26
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   439
      TabIndex        =   0
      Top             =   0
      Width           =   6585
   End
End
Attribute VB_Name = "JCToolbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==========================================================================
'   Copyright © 2005 Juan Carlos San Román Arias
'
'   JC_ToolBar
'   Data: 23-Nov-2005
'
'   This is an Office 2003 toolbar for VB. You can built a nice Toolbar.
'   The initial idea taken from JCF_Toolbutton created by João Fortes.
'   I have made a compilation of different jobs published on Planet-Source-Code.com
'   I want to thank to
'   - Everyday Panos for your Office 2003 Button AND MOVING TOOLBAR project
'   - Fred cpp for api functions used in his isbutton control
'   - Carles P.V. for 3d UcVertical line
'   - All control is drawn using api functions (no images, no other controls)
'   Data: 23/11/2005
'
'==========================================================================
Option Explicit

'state constants
Const STA_NORMAL = 0
Const STA_OVER = 1
Const STA_DOWN = 2
Const STA_OVERDOWN = 3
Const STA_DISABLED = 4

'xp theme
Public Enum ThemeConst
    Blue = 0
    Silver = 1
    Olive = 2
    Custom = 3
    Autodetect = 4
End Enum

'events
Event ButtonClick()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

'members
Dim m_IsStrech As Boolean
Dim m_ThemeColor As ThemeConst

'local
Dim tmpState As Integer

Private ColorFrom As OLE_COLOR, ColorTo As OLE_COLOR
Private ColorToolbar As OLE_COLOR, ColorBorderPic As OLE_COLOR
Private ColorToRight As OLE_COLOR, ColorFromRight As OLE_COLOR

Private Sub PicRight_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    tmpState = STA_DOWN
    DrawRight RGB(255, 154, 87), RGB(255, 212, 144)
End Sub

Private Sub PicRight_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If tmrRight.Enabled = False Then tmrRight.Enabled = True
End Sub

Private Sub PicRight_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    tmpState = STA_OVER
    DrawRight RGB(255, 244, 204), RGB(255, 197, 125)
End Sub

Private Sub tmrRight_Timer()
    If CheckMouseOver Then
        If tmpState = STA_NORMAL Then
            tmpState = STA_OVER
            DrawRight RGB(255, 244, 204), RGB(255, 197, 125)
        End If
    Else
        If tmpState = STA_OVER Or tmpState = STA_DOWN Then
            tmpState = STA_NORMAL
        End If
        DrawRight ColorFromRight, ColorToRight
    End If
End Sub

'==========================================================================
' Init, Read & Write UserControl
'==========================================================================

Private Sub UserControl_Initialize()
    'Color selection according to window setup
    m_ThemeColor = Autodetect
    Call SetThemeColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
End Sub

Private Sub UserControl_InitProperties()
    UserControl.BackColor = Ambient.BackColor
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 390
    PicGrad.Move PicLeft.Left + PicLeft.Width, 0, UserControl.ScaleWidth - PicLeft.Width - PicRight.Width, UserControl.ScaleHeight
    PicRight.Left = PicGrad.Left + PicGrad.Width
    DrawTheme
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_ThemeColor = PropBag.ReadProperty("ThemeColor", Autodetect)
    m_IsStrech = PropBag.ReadProperty("IsStrech", False)
    PicRight.BackColor = UserControl.BackColor
    PicLeft.BackColor = UserControl.BackColor
    Call SetThemeColor
    DrawTheme
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("ThemeColor", m_ThemeColor, Autodetect)
    Call PropBag.WriteProperty("IsStrech", m_IsStrech, False)
End Sub
'==========================================================================
' Down
'==========================================================================
Private Sub picleft_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

Private Sub picleft_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, x, Y)
End Sub
'==========================================================================
' Move
'==========================================================================
Private Sub picleft_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub
'==========================================================================
' Click
'==========================================================================
Private Sub PicRight_Click()
    RaiseEvent ButtonClick
End Sub

'==========================================================================
' Properties
'==========================================================================

Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor = New_BackColor
    PicRight.BackColor = New_BackColor
    PicLeft.BackColor = New_BackColor
    PropertyChanged "BackColor"
    Call DrawTheme
End Property

Public Property Get IsStrech() As Boolean
    IsStrech = m_IsStrech
End Property

Public Property Let IsStrech(ByVal New_Value As Boolean)
    m_IsStrech = New_Value
    DrawRight ColorFromRight, ColorToRight
    PropertyChanged "IsStrech"
End Property

Public Property Get ThemeColor() As ThemeConst
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As ThemeConst)
    m_ThemeColor = vData
    Call SetThemeColor
    DrawTheme
    PropertyChanged "ThemeColor"
End Property

'==========================================================================
' Functions
'==========================================================================
Public Sub DrawTheme()
    Dim R As RECT
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    DrawVGradientEx PicGrad.hdc, ColorTo, ColorFrom, R.Left, R.Top, R.Right, R.Bottom
    SetRect R, 0, UserControl.ScaleHeight - 1, UserControl.ScaleWidth, UserControl.ScaleHeight - 1
    APILineEx PicGrad.hdc, R.Left, R.Top, R.Right, R.Bottom, ColorToolbar
    PicGrad.Refresh
    DrawLeft
    DrawRight ColorFromRight, ColorToRight
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINT
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.x, pt.Y) = PicRight.hwnd)
    tmrRight.Enabled = CheckMouseOver
End Function

Private Sub DrawRight(FromColor As Long, ToColor As Long)
    Dim R As RECT, lcolor As Long
    Dim poly(1 To 3) As POINT, i As Integer
    
    SetRect R, 0, 0, 2, PicRight.Height
    DrawVGradientEx PicRight.hdc, ColorTo, ColorFrom, R.Left, R.Top, R.Right, R.Bottom
    
    SetRect R, 2, 0, PicRight.Width, PicRight.Height
    DrawVGradientEx PicRight.hdc, FromColor, ToColor, R.Left, R.Top, R.Right, R.Bottom
    
    lcolor = TranslateColor(Ambient.BackColor)
    SetPixel PicRight.hdc, PicRight.Width - 1, 0, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, PicRight.Height - 1, lcolor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), FromColor)
    SetPixel PicRight.hdc, PicRight.Width - 2, 0, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, 1, lcolor
    
    SetPixel PicRight.hdc, 0, 0, lcolor
    SetPixel PicRight.hdc, 1, 1, lcolor
    SetPixel PicRight.hdc, 1, 0, FromColor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), ToColor)
    SetPixel PicRight.hdc, PicRight.Width - 2, PicRight.Height - 1, lcolor
    SetPixel PicRight.hdc, PicRight.Width - 1, PicRight.Height - 2, lcolor
    
    SetPixel PicRight.hdc, 0, PicRight.Height - 1, lcolor
    SetPixel PicRight.hdc, 1, PicRight.Height - 2, lcolor
    SetPixel PicRight.hdc, 1, PicRight.Height - 1, ToColor
    
    'drawing big right arrow
        'white triangle
        poly(1).x = 7:  poly(1).Y = 19
        poly(2).x = 7 + 6: poly(2).Y = 19
        poly(3).x = 7 + 3: poly(3).Y = 19 + 3
        DrawTriangle PicRight, vbWhite, WHITE_BRUSH, poly, 3
        
        'black triangle
        poly(1).x = 6:  poly(1).Y = 18
        poly(2).x = 6 + 6: poly(2).Y = 18
        poly(3).x = 6 + 3: poly(3).Y = 18 + 3
        DrawTriangle PicRight, vbBlack, BLACKBRUSH, poly, 3
        
        'black line
        SetRect R, 6, 15, 13, 15
        APILineEx PicRight.hdc, R.Left, R.Top, R.Right, R.Bottom, vbBlack
        
        'white line
        SetRect R, 7, 16, 14, 16
        APILineEx PicRight.hdc, R.Left, R.Top, R.Right, R.Bottom, vbWhite
    
    If m_IsStrech Then
        'drawing small arrows
        For i = 0 To 1
            SetRect R, 6 + 4 * i, 5, 6 + 4 * i, 8
            APILineEx PicRight.hdc, R.Left, R.Top, R.Right, R.Bottom, vbBlack
            SetPixel PicRight.hdc, 7 + 4 * i, 6, vbBlack
            SetPixel PicRight.hdc, 7 + 4 * i, 7, vbWhite
            SetPixel PicRight.hdc, 8 + 4 * i, 7, vbWhite
            SetPixel PicRight.hdc, 7 + 4 * i, 8, vbWhite
        Next i
    End If
   
    PicRight.Refresh
End Sub

Private Sub DrawLeft()
    Dim R As RECT, lcolor As Long, i As Long
    
    SetRect R, 0, 0, PicLeft.Width, PicLeft.Height
    DrawVGradientEx PicLeft.hdc, ColorTo, ColorFrom, R.Left, R.Top, R.Right, R.Bottom

    SetRect R, 2, PicLeft.Height - 1, PicLeft.Width, PicLeft.Height - 1
    APILineEx PicLeft.hdc, R.Left, R.Top, R.Right, R.Bottom, ColorToolbar

    lcolor = TranslateColor(Ambient.BackColor)
    SetPixel PicLeft.hdc, 0, 0, lcolor
    SetPixel PicLeft.hdc, 0, PicRight.Height - 1, lcolor
    SetPixel PicLeft.hdc, 0, PicRight.Height - 2, lcolor
    SetPixel PicLeft.hdc, 1, PicRight.Height - 1, lcolor

    lcolor = BlendColors(vbWhite, ColorTo)
    SetPixel PicLeft.hdc, 1, 0, lcolor
    SetPixel PicLeft.hdc, 0, 1, lcolor

    lcolor = BlendColors(ColorBorderPic, ColorFrom)
    SetPixel PicLeft.hdc, 1, PicRight.Height - 3, lcolor
    
    lcolor = BlendColors(TranslateColor(Ambient.BackColor), ColorFrom)
    SetPixel PicLeft.hdc, 0, PicRight.Height - 3, lcolor
    SetPixel PicLeft.hdc, 1, PicRight.Height - 2, lcolor

    For i = 0 To 3
        SetRect R, 5, 5 + 4 * i, 1, 1
        APIRectangle PicLeft.hdc, R.Left, R.Top, R.Right, R.Bottom, vbWhite
        
        SetRect R, 4, 4 + 4 * i, 1, 1
        APIRectangle PicLeft.hdc, R.Left, R.Top, R.Right, R.Bottom, ColorBorderPic
    Next i
    PicLeft.Refresh
End Sub

Private Sub SetThemeColor()
    
    Select Case m_ThemeColor
        Case Is = Autodetect
            Call GetGradientColor(UserControl.hwnd)
        Case Else
            SetDefaultThemeColor m_ThemeColor
    End Select
End Sub

Private Sub SetDefaultThemeColor(ThemeType As Long)
    Select Case ThemeType
            Case 0 '"NormalColor"
                ColorFrom = RGB(129, 169, 226)
                ColorTo = RGB(221, 236, 254)
                ColorToolbar = RGB(59, 97, 156)
                ColorBorderPic = RGB(0, 0, 128)
                ColorFromRight = RGB(118, 167, 241)
                ColorToRight = RGB(0, 53, 145)
            Case 1 '"Metallic"
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
            Case 2 '"HomeStead"
                ColorFrom = RGB(181, 197, 143)
                ColorTo = RGB(247, 249, 225)
                ColorToolbar = RGB(96, 128, 88)
                ColorBorderPic = RGB(63, 93, 56)
                ColorFromRight = RGB(177, 195, 141)
                ColorToRight = RGB(96, 119, 107)
            Case 3 '"Custom"
                ColorFrom = RGB(181, 197, 143)
                ColorTo = RGB(247, 249, 225)
                ColorToolbar = RGB(96, 128, 88)
                ColorBorderPic = RGB(63, 93, 56)
                ColorFromRight = RGB(177, 195, 141)
                ColorToRight = RGB(96, 119, 107)
            Case Else
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
        End Select
End Sub

Private Sub GetGradientColor(lhWnd As Long)

    GetThemeName lhWnd
    
    If AppThemed Then   '/Check if themed.
        Select Case m_sCurrentSystemThemename
            Case "NormalColor"
                ColorFrom = RGB(129, 169, 226)
                ColorTo = RGB(221, 236, 254)
                ColorToolbar = RGB(59, 97, 156)
                ColorBorderPic = RGB(0, 0, 128)
                ColorFromRight = RGB(118, 167, 241)
                ColorToRight = RGB(0, 53, 145)
            Case "Metallic"
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
            Case "HomeStead"
                ColorFrom = RGB(181, 197, 143)
                ColorTo = RGB(247, 249, 225)
                ColorToolbar = RGB(96, 128, 88)
                ColorBorderPic = RGB(63, 93, 56)
                ColorFromRight = RGB(177, 195, 141)
                ColorToRight = RGB(96, 119, 107)
            Case Else
                ColorFrom = RGB(153, 151, 180)
                ColorTo = RGB(244, 244, 251)
                ColorToolbar = RGB(124, 124, 148)
                ColorBorderPic = RGB(75, 75, 111)
                ColorFromRight = RGB(180, 179, 200)
                ColorToRight = RGB(118, 116, 146)
        End Select
    Else
'        glColorOneNormal = BlendColor(vbButtonFace, vbWhite, 120)
'        glColorTwoNormal = vbButtonFace
'        glColorBorder = BlendColor(vbButtonFace, vbBlack, 200)
'        glColorHeaderColorOne = BlendColor(vbButtonFace, vbWhite, 120)
'        glColorHeaderColorTwo = vbButtonFace
'        glColorOneSelected = BlendColor(vbInactiveTitleBar, BlendColor(vbButtonFace, vbWhite, 150), 100)
'        glColorTwoSelected = glColorOneSelected
        ColorFrom = RGB(153, 151, 180)
        ColorTo = RGB(244, 244, 251)
        ColorToolbar = RGB(124, 124, 148)
        ColorBorderPic = RGB(75, 75, 111)
        ColorFromRight = RGB(180, 179, 200)
        ColorToRight = RGB(118, 116, 146)
    End If
End Sub



