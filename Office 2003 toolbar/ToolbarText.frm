VERSION 5.00
Object = "{92C3F590-2B79-4E50-9664-E1CDAC443CA0}#5.0#0"; "jcOffice2003.ocx"
Begin VB.Form ToolbarText 
   Caption         =   "Office 2003 style toolbar and toolbar button"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5655
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "JCF_Button improving features (created by João Fortes):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   120
      TabIndex        =   29
      Top             =   2910
      Width           =   6105
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "aaaa"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   330
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Office 2003 style toolbar control:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   120
      TabIndex        =   27
      Top             =   1290
      Width           =   6105
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "aaa"
         Height          =   195
         Left            =   180
         TabIndex        =   28
         Top             =   300
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Apply states on Help button:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Width           =   6105
      Begin VB.CommandButton Command1 
         Caption         =   "Normal"
         Height          =   390
         Index           =   0
         Left            =   240
         TabIndex        =   26
         Top             =   330
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Over"
         Height          =   390
         Index           =   1
         Left            =   1395
         TabIndex        =   25
         Top             =   330
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Down"
         Height          =   390
         Index           =   2
         Left            =   2550
         TabIndex        =   24
         Top             =   330
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "OverDown"
         Height          =   390
         Index           =   3
         Left            =   3705
         TabIndex        =   23
         Top             =   330
         Width           =   990
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Disabled"
         Height          =   390
         Index           =   4
         Left            =   4875
         TabIndex        =   22
         Top             =   330
         Width           =   990
      End
   End
   Begin jcOffice2003.JCToolbar JCToolbar1 
      Height          =   390
      Index           =   2
      Left            =   0
      TabIndex        =   18
      Top             =   780
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   688
      BackColor       =   -2147483633
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   10
         Left            =   210
         TabIndex        =   20
         ToolTipText     =   "You can put text without icon"
         Top             =   0
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         Caption         =   "Only text"
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   11
         Left            =   1260
         TabIndex        =   19
         ToolTipText     =   "You can put icon without text"
         Top             =   0
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   661
         Picture         =   "ToolbarText.frx":0000
      End
   End
   Begin jcOffice2003.JCToolbar JCToolbar1 
      Height          =   390
      Index           =   1
      Left            =   0
      TabIndex        =   1
      Top             =   390
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   688
      BackColor       =   -2147483633
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   9
         Left            =   5190
         TabIndex        =   17
         ToolTipText     =   "Exit this demo"
         Top             =   0
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         Caption         =   "Exit"
         Picture         =   "ToolbarText.frx":05AA
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   5
         Left            =   3720
         TabIndex        =   16
         Top             =   60
         Width           =   60
         _ExtentX        =   159
         _ExtentY        =   450
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   15
         Top             =   60
         Width           =   60
         _ExtentX        =   159
         _ExtentY        =   450
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   8
         Left            =   3900
         TabIndex        =   14
         ToolTipText     =   "Full screen viewing"
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         Caption         =   "Full screen"
         Picture         =   "ToolbarText.frx":0B44
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   7
         Left            =   2880
         TabIndex        =   13
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Find"
         Picture         =   "ToolbarText.frx":10DE
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   6
         Left            =   1830
         TabIndex        =   11
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Send"
         Picture         =   "ToolbarText.frx":1678
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   5
         Left            =   210
         TabIndex        =   10
         ToolTipText     =   "Dropdown button for opening lisbox or popup menu"
         Top             =   0
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         Caption         =   "Shortcuts"
         IsCheckButton   =   -1  'True
         IsDropDown      =   -1  'True
         Picture         =   "ToolbarText.frx":1C12
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   3
         Left            =   1710
         TabIndex        =   6
         Top             =   60
         Width           =   60
         _ExtentX        =   159
         _ExtentY        =   450
      End
   End
   Begin jcOffice2003.JCToolbar JCToolbar1 
      Height          =   390
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   688
      BackColor       =   -2147483633
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   4
         Left            =   4380
         TabIndex        =   12
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Help"
         Picture         =   "ToolbarText.frx":21BC
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   3
         Left            =   3360
         TabIndex        =   9
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Undo"
         Picture         =   "ToolbarText.frx":2756
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   2
         Left            =   2280
         TabIndex        =   8
         ToolTipText     =   "Example of check button"
         Top             =   0
         Width           =   765
         _ExtentX        =   1349
         _ExtentY        =   661
         Caption         =   "Left"
         IsCheckButton   =   -1  'True
         Picture         =   "ToolbarText.frx":2CF0
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         ToolTipText     =   "Example of check button"
         Top             =   0
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   661
         Caption         =   "Right"
         IsCheckButton   =   -1  'True
         Picture         =   "ToolbarText.frx":328A
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   2
         Left            =   4290
         TabIndex        =   5
         Top             =   60
         Width           =   60
         _ExtentX        =   159
         _ExtentY        =   450
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   1
         Left            =   3180
         TabIndex        =   4
         Top             =   60
         Width           =   60
         _ExtentX        =   159
         _ExtentY        =   450
      End
      Begin jcOffice2003.ucVertical3DLine Line3d 
         Height          =   240
         Index           =   0
         Left            =   1140
         TabIndex        =   3
         Top             =   60
         Width           =   120
         _ExtentX        =   159
         _ExtentY        =   450
      End
      Begin jcOffice2003.JCF_Button btn 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   2
         Top             =   0
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
         Caption         =   "Home"
         Picture         =   "ToolbarText.frx":3824
      End
   End
   Begin VB.Timer TimMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   7080
      Top             =   2640
   End
   Begin VB.Menu mnufile 
      Caption         =   "file"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu mnuoption 
         Caption         =   "Hide Toolbar"
         Index           =   0
      End
      Begin VB.Menu mnuoption 
         Caption         =   "Show hidden toolbars"
         Index           =   1
      End
      Begin VB.Menu mnuoption 
         Caption         =   "Change Theme color"
         Index           =   2
         Begin VB.Menu MnuTheme 
            Caption         =   "Blue"
            Index           =   0
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Silver"
            Index           =   1
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Olive"
            Index           =   2
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Custom"
            Index           =   3
         End
         Begin VB.Menu MnuTheme 
            Caption         =   "Autodetect"
            Index           =   4
         End
      End
   End
End
Attribute VB_Name = "ToolbarText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NeoX As Long, IniWidth() As Integer, selection As Integer, blnhide As Boolean

Private Sub Command1_Click(Index As Integer)
    btn(4).State = Index
End Sub

Private Sub btn_Click(Index As Integer)

    If Index = 1 Then
        If btn(2).Value = 0 Then
            btn(2).Value = 1
        Else
            btn(2).Value = 0
        End If
    End If

    If Index = 2 Then
        If btn(1).Value = 0 Then
            btn(1).Value = 1
        Else
            btn(1).Value = 0
        End If
    End If

    If Index = 9 Then
        Unload Me
    Else
        If btn(Index).Caption <> "" Then
            MsgBox """" & btn(Index).Caption & """ button pressed"
        Else
            MsgBox "btn(" & Index & ") button pressed"
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    btn(1).Value = 1
    ReDim IniWidth(JCToolbar1.Count)
    For i = 0 To JCToolbar1.Count - 1
        IniWidth(i) = JCToolbar1(i).Width
        JCToolbar1(i).ThemeColor = i
        ChangeThemeButton btn, JCToolbar1(i), i
    Next i
    
    Label3.Caption = "- It can be moved horizontally" & Chr(13)
    Label3.Caption = Label3.Caption & "- It resizes when form is resized " & Chr(13)
    Label3.Caption = Label3.Caption & "- Control is fully drawn using api functions " & Chr(13)
    Label3.Caption = Label3.Caption & "- windows XP theme auto detection or selection  (blue, silver and olive) " & Chr(13)
    Label3.Caption = Label3.Caption & "- There is a buttonclick event for toolbar right side "

    Label4.Caption = "- The way of putting icon image and text on the button has changed" & Chr(13)
    Label4.Caption = Label4.Caption & "- Added dropdown feature, useful for popup menu and lisbox using" & Chr(13)
    Label4.Caption = Label4.Caption & "- windows XP theme auto detection or selection  (blue, silver and olive) " & Chr(13)
    Label4.Caption = Label4.Caption & "- You can put only icon or text" & Chr(13)
    Label4.Caption = Label4.Caption & "- Added function to convert color icon to grayscale icon when button is disabled"
End Sub

Private Sub Form_Resize()
    Dim i As Integer
    
    If (Me.WindowState = vbMinimized) Then
        Me.Visible = False
        bln_minimized = True
        Exit Sub
    Else    'not minimized
        For i = 0 To JCToolbar1.Count - 1
            ResizeToolbar JCToolbar1(i), btn, IniWidth(i), Line3d
        Next i
    End If
End Sub

Private Sub JCToolbar1_ButtonClick(Index As Integer)
    selection = Index
    Me.PopupMenu mnufile(0), vbPopupMenuRightAlign, JCToolbar1(Index).Left + JCToolbar1(Index).Width, JCToolbar1(Index).Top + JCToolbar1(Index).Height
End Sub

Private Sub JCToolbar1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        TimMove.Enabled = True
    End If
    selection = Index
End Sub

Private Sub JCToolbar1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NeoX = JCToolbar1(Index).Left + 15 * X
    If NeoX < 0 Then
        NeoX = 0
    ElseIf NeoX > Me.Width - JCToolbar1(Index).Width Then
        NeoX = Me.Width - JCToolbar1(Index).Width - 100
    End If
End Sub

Private Sub JCToolbar1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimMove.Enabled = False
End Sub

Private Sub mnuoption_Click(Index As Integer)
    Dim i As Integer
    Select Case Index
        Case 0  'hide a toolbar
            JCToolbar1(selection).Visible = False
        Case 1 'show hidden toolbars
            For i = 0 To JCToolbar1.Count - 1
                If JCToolbar1(i).Visible = False Then JCToolbar1(i).Visible = True
            Next i
    End Select
End Sub

Private Sub MnuTheme_Click(Index As Integer)
    JCToolbar1(selection).ThemeColor = Index
    ChangeThemeButton btn, JCToolbar1(selection), Index
End Sub

Private Sub TimMove_Timer()
    JCToolbar1(selection).Left = NeoX
End Sub

Private Sub ResizeToolbar(ToolBar As Object, BtnA As Object, IniWidth As Integer, Line3 As Object)
    blnhide = False
    
    If Me.ScaleWidth < (ToolBar.Left + ToolBar.Width) Then
        If (Me.ScaleWidth - ToolBar.Left) > 450 Then
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
        CheckControl BtnA, ToolBar
        CheckControl Line3, ToolBar
    Else
        If (ToolBar.Left + IniWidth) < Me.ScaleWidth Then
            ToolBar.Width = IniWidth
        Else
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
        CheckControl BtnA, ToolBar
        CheckControl Line3, ToolBar
    End If
    ToolBar.IsStrech = blnhide
End Sub

Private Sub ChangeThemeButton(BtnA As Object, cont As Control, Theme As Integer)
    Dim i As Integer
    For i = 0 To BtnA.Count - 1
        If BtnA(i).Container Is cont Then
            BtnA(i).ThemeColor = Theme
        End If
    Next i
End Sub

Public Sub CheckControl(Ctrl As Object, ToolBar As Object)
    Dim i As Integer
    For i = 0 To Ctrl.Count - 1
        If Ctrl(i).Container Is ToolBar Then
            If ToolBar.Width - 210 < Ctrl(i).Left + Ctrl(i).Width Then
                Ctrl(i).Visible = False
                blnhide = True
            Else
                Ctrl(i).Visible = True
            End If
        End If
    Next i
End Sub
