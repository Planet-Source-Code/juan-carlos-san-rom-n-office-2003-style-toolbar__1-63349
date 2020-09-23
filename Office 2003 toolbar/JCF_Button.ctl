VERSION 5.00
Begin VB.UserControl JCF_Button 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   360
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1380
   ScaleHeight     =   24
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   92
   ToolboxBitmap   =   "JCF_Button.ctx":0000
   Begin VB.Timer tmrOver 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   90
      Top             =   930
   End
   Begin VB.PictureBox PicMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   0
      ScaleHeight     =   23
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   0
      Top             =   0
      Width           =   1395
   End
   Begin VB.Image ImgArrow 
      Height          =   240
      Left            =   1530
      Picture         =   "JCF_Button.ctx":0312
      Top             =   30
      Width           =   240
   End
End
Attribute VB_Name = "JCF_Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'==========================================================================
'   This code improves JCF_ToolButton created by João Fortes
'
'   You can built a nice toolbar button. I Just improve and join same other code peaces from PSC. THANKS PSC
'
'   - I added grayscale icon effect for disabled property used in the modification of
'     Fred cpp isbutton control 3.0.
'
'   I want to thank to
'   - Everyday Panos for your Office 2003 Button AND MOVING TOOLBAR project
'   - Fred cpp for api functions used in his isbutton control
'
'   Data: 23/11/2005
'
'==========================================================================
Option Explicit

'API contants
Private Const DSS_DISABLED = &H20
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10
Private Const DST_BITMAP = &H4
Private Const DST_ICON = &H3
Private Const DST_PREFIXTEXT = &H2
Private Const DST_TEXT = &H1

'constants
Const TEXT_ACTIVE = &H0&
Const TEXT_INACTIVE = &H6A6A6A '&H80000011          '&H6A6A6A

'state constants
Const STA_NORMAL = 0
Const STA_OVER = 1
Const STA_DOWN = 2
Const STA_OVERDOWN = 3
Const STA_DISABLED = 4

'value constants
Const VAL_UNCHECKED = 0
Const VAL_CHECKED = 1
Const VAL_GRAY = 2

'aligment contants
Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0

'xp theme
Public Enum ThemeConstA
    Blue = 0
    Silver = 1
    Olive = 2
    Custom = 3
    Autodetect = 4
End Enum

'events
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

'members
Dim m_Icon As StdPicture
Dim m_IconSize As Long
Dim m_State As Integer
Dim m_Value As Integer
Dim m_Enabled As Boolean
Dim m_Caption As String
Dim m_IsCheckButton As Boolean
Dim m_IsDropDown As Boolean
Dim m_CaptionColor As Long
Dim m_ThemeColor As ThemeConstA

'local
Dim tmpState As Integer
Dim tmpDrawState As Integer

Private ColorFrom As OLE_COLOR, ColorTo As OLE_COLOR
Private ColorToolbar As OLE_COLOR, ColorBorderPic As OLE_COLOR
Private ColorToRight As OLE_COLOR, ColorFromRight As OLE_COLOR

'==========================================================================
' Init, Read & Write UserControl
'==========================================================================
Private Sub UserControl_InitProperties()
    m_Enabled = True
    m_Caption = Ambient.DisplayName
End Sub

Private Sub UserControl_Initialize()
    tmpDrawState = -1
    m_ThemeColor = Autodetect
    Call SetThemeColor
End Sub

Private Sub UserControl_Resize()
    UserControl.Height = 375
    PicMain.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    Call DrawButton
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        m_State = .ReadProperty("State", 0)
        m_Enabled = .ReadProperty("Enabled", True)
        m_ThemeColor = .ReadProperty("ThemeColor", Autodetect)
        m_Caption = .ReadProperty("Caption", Empty)
        m_IsCheckButton = .ReadProperty("IsCheckButton", False)
        m_IsDropDown = .ReadProperty("IsDropDown", False)
        Set m_Icon = .ReadProperty("Picture", Nothing)
    End With
    Call SetThemeColor
    Call DrawButton
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("State", m_State, 0)
        Call .WriteProperty("Enabled", m_Enabled, True)
        Call .WriteProperty("ThemeColor", m_ThemeColor, Autodetect)
        Call .WriteProperty("Caption", m_Caption, Empty)
        Call .WriteProperty("IsCheckButton", m_IsCheckButton, False)
        Call .WriteProperty("IsDropDown", m_IsDropDown, False)
        Call .WriteProperty("Picture", m_Icon, Nothing)
    End With
End Sub
'==========================================================================
' Down
'==========================================================================
Private Sub PicMain_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    
    'only LeftButton
    If Button <> vbLeftButton Then Exit Sub
    
    tmpState = m_State
    m_State = STA_DOWN
    
   ' DrawButton
    If m_IsCheckButton Then
        Call DrawStateBtn(RGB(254, 214, 145), RGB(254, 142, 75), 0) 'down
        m_CaptionColor = TEXT_ACTIVE = TEXT_ACTIVE
    End If

    RaiseEvent MouseDown(Button, Shift, x, Y)
End Sub

'==========================================================================
' Up
'==========================================================================
Private Sub PicMain_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    
    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    
    'type of button
    If m_IsCheckButton Then
        If m_Value = VAL_UNCHECKED Then
            m_Value = VAL_CHECKED
        ElseIf m_Value = VAL_CHECKED Then
            m_Value = VAL_UNCHECKED
        End If
    Else
        m_State = STA_NORMAL
    End If
    
    'Fire event
    RaiseEvent MouseUp(Button, Shift, x, Y)
    RaiseEvent Click
End Sub

'==========================================================================
' Move
'==========================================================================
Private Sub PicMain_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
    'Is disabled
    If m_State = STA_DISABLED Or Not m_Enabled Then Exit Sub
    
    PicMain.ToolTipText = Extender.ToolTipText

    If tmrOver.Enabled = False Then tmrOver.Enabled = True
    RaiseEvent MouseMove(Button, Shift, x, Y)
End Sub

'==========================================================================
' MouseOver Status
'==========================================================================
Private Sub tmrOver_Timer()
    If CheckMouseOver Then
        If m_IsCheckButton Then
            If m_Value = VAL_UNCHECKED Then
                m_State = STA_OVER
            ElseIf m_Value = VAL_CHECKED Then
                If m_IsDropDown Then
                    m_State = STA_DOWN
                Else
                    m_State = STA_OVERDOWN
                End If
            End If
            DrawButton True, True
        Else
            If m_State = STA_NORMAL Then
                m_State = STA_OVER
            End If
            DrawButton True, True
        End If
    Else
        If m_State = STA_OVER Then
            m_State = STA_NORMAL
        ElseIf m_State = STA_OVERDOWN Then
            m_State = STA_DOWN
        End If

        DrawButton False, True
    End If
End Sub

'==========================================================================
' Properties
'==========================================================================
Public Property Get Value() As Integer
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Integer)
    m_Value = New_Value
    DrawButton
    PropertyChanged "Value"
End Property

Public Property Get IsCheckButton() As Boolean
    IsCheckButton = m_IsCheckButton
End Property

Public Property Let IsCheckButton(ByVal New_Value As Boolean)
    m_IsCheckButton = New_Value
    DrawButton
    PropertyChanged "IsCheckButton"
End Property

Public Property Get IsDropDown() As Boolean
    IsDropDown = m_IsDropDown
End Property

Public Property Let IsDropDown(ByVal New_Value As Boolean)
    m_IsDropDown = New_Value
    DrawButton
    PropertyChanged "IsDropDown"
End Property

Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Value As String)
    m_Caption = New_Value
    DrawButton
    PropertyChanged "Caption"
End Property

Public Property Get State() As Integer
    State = m_State
End Property

Public Property Let State(ByVal New_Value As Integer)
    m_State = New_Value
    DrawButton
    PropertyChanged "State"
End Property

Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Value As Boolean)
    m_Enabled = New_Value
    PropertyChanged "Enabled"
End Property

Public Property Get Picture() As StdPicture
    Set Picture = m_Icon
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
    Set m_Icon = New_Picture
    DrawButton
    PropertyChanged "Picture"
End Property

Public Property Get ThemeColor() As ThemeConstA
    ThemeColor = m_ThemeColor
End Property

Public Property Let ThemeColor(ByVal vData As ThemeConstA)
    m_ThemeColor = vData
    Call SetThemeColor
    'Is disabled
    If m_State = STA_DISABLED Then
        Call DrawStateBtn(ColorFrom, ColorTo, 3)
    Else
        DrawButton
    End If
    PropertyChanged "ThemeColor"
End Property
'==========================================================================
' Functions
'==========================================================================
Public Sub DrawButton(Optional IsOver As Boolean = False, Optional FromTimer As Boolean = False)

    If m_State <> STA_DISABLED And FromTimer = False Then
        If m_IsCheckButton Then
            If IsOver Then
                If m_Value = VAL_CHECKED Then
                    If m_IsDropDown Then
                        m_State = STA_DOWN
                    Else
                        m_State = STA_OVERDOWN
                    End If
                Else
                    m_State = STA_OVER
                End If
            Else
                If m_Value = VAL_CHECKED Then
                    m_State = STA_DOWN
                Else
                    m_State = STA_NORMAL
                End If
            End If
        End If
    End If
    
    'No need to redraw it
    If tmpDrawState > 0 And m_State = tmpDrawState Then Exit Sub
    
    m_CaptionColor = TEXT_ACTIVE
    
    Select Case m_State
        Case STA_NORMAL
            Call DrawStateBtn(ColorFrom, ColorTo, 1)  'normal
        Case STA_OVER
            Call DrawStateBtn(RGB(255, 207, 142), RGB(255, 245, 206), 0) 'over
        Call DrawStateBtn(RGB(255, 207, 142), RGB(255, 245, 206), 0) 'over
        Case STA_DOWN
            If m_IsDropDown Then
                Call DrawStateBtn(ColorFrom, RGB(255, 255, 255), 2) 'downdrop
            Else
                Call DrawStateBtn(RGB(254, 214, 145), RGB(254, 142, 75), 0) 'down
            End If
        Case STA_OVERDOWN
            Call DrawStateBtn(RGB(255, 119, 0), RGB(255, 235, 99), 0) 'overdown
        Case STA_DISABLED
            m_CaptionColor = TEXT_INACTIVE
            Call DrawStateBtn(ColorFrom, ColorTo, 3)  'normal
    End Select

    tmpDrawState = m_State
End Sub

Private Function CheckMouseOver() As Boolean
    Dim pt As POINT
    GetCursorPos pt
    CheckMouseOver = (WindowFromPoint(pt.x, pt.Y) = PicMain.hwnd)
    tmrOver.Enabled = CheckMouseOver
End Function

Private Sub DrawStateBtn(ColorFromA As Long, ColorToA As Long, intoption As Integer)
    Dim R As RECT
    m_IconSize = 16
    
    'drawing background
    SetRect R, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    DrawVGradientEx PicMain.hdc, ColorTo, ColorFrom, R.Left, R.Top, R.Right, R.Bottom
    
    'drawing rectangle
    Select Case intoption
        Case 0
            SetRect R, 0, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 5
            DrawGradBorderRect PicMain.hdc, ColorToA, ColorFromA, R, ColorBorderPic
        Case 2
            SetRect R, 0, 2, UserControl.ScaleWidth - 1, UserControl.ScaleHeight + 1
            DrawGradBorderRect PicMain.hdc, ColorFromA, ColorToA, R, ColorBorderPic
    End Select
    
    'drawing image
    If Not (m_Icon Is Nothing) Then
        If intoption <> 3 Then
           PicMain.PaintPicture m_Icon, 5, 5, m_IconSize, m_IconSize
        Else
           PaintIconGrayscale PicMain.hdc, m_Icon, 5, 5, m_IconSize, m_IconSize
        End If
    End If
    
    'drawing caption
    If Not (m_Caption = "") Then DrawCaption m_Caption, m_CaptionColor, 1

    'drawing arrow
    If m_IsDropDown Then PicMain.PaintPicture ImgArrow.Picture, UserControl.ScaleWidth - ImgArrow.Width, (UserControl.ScaleHeight - ImgArrow.Height) / 2
    
    PicMain.Refresh
End Sub

Private Sub DrawCaption(strText As String, FntColor As Long, XYOffset As Integer)
    Dim R As RECT
    
    PicMain.ForeColor = FntColor
    
    'Set the rectangle's values
    If m_Icon Is Nothing Then
        SetRect R, 5, 0, PicMain.ScaleWidth, PicMain.ScaleHeight
    Else
        SetRect R, 25, 0, PicMain.ScaleWidth, PicMain.ScaleHeight
    End If

    'Draw text in PicMain
    DrawTextEx PicMain.hdc, strText, Len(strText), R, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, ByVal 0&
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


