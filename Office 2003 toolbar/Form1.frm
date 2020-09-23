VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   2310
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin Project1.JCToolbar JCToolbar1 
      Height          =   390
      Index           =   1
      Left            =   0
      TabIndex        =   11
      Top             =   390
      Width           =   6765
      _extentx        =   11933
      _extenty        =   688
      backcolor       =   -2147483633
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   8
         Left            =   5370
         TabIndex        =   19
         Top             =   0
         Width           =   1035
         _extentx        =   1826
         _extenty        =   661
         caption         =   "&Proving"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":0000
      End
      Begin Project1.ucVertical3DLine Line3d 
         Height          =   255
         Index           =   2
         Left            =   2190
         TabIndex        =   18
         Top             =   60
         Width           =   90
         _extentx        =   159
         _extenty        =   450
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   7
         Left            =   3930
         TabIndex        =   15
         Top             =   0
         Width           =   1035
         _extentx        =   1826
         _extenty        =   661
         caption         =   "&Justified"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":019E
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   6
         Left            =   180
         TabIndex        =   14
         ToolTipText     =   "Enviar por correio electrónico"
         Top             =   0
         Width           =   840
         _extentx        =   1482
         _extenty        =   661
         caption         =   "&Send"
         picture         =   "Form1.frx":033C
         picturedisabled =   "Form1.frx":08D6
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   5
         Left            =   1170
         TabIndex        =   13
         Top             =   0
         Width           =   990
         _extentx        =   1746
         _extenty        =   661
         caption         =   "&Portrait"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":0A80
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   4
         Left            =   2280
         TabIndex        =   12
         Top             =   0
         Width           =   1395
         _extentx        =   2461
         _extenty        =   661
         caption         =   "&Landscape"
         ischeckbutton   =   -1  'True
         isdropdown      =   -1  'True
         picture         =   "Form1.frx":101A
      End
   End
   Begin VB.Timer TimMove 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6030
      Top             =   1110
   End
   Begin Project1.JCToolbar JCToolbar1 
      Height          =   390
      Index           =   0
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5265
      _extentx        =   9287
      _extenty        =   688
      backcolor       =   -2147483633
      Begin Project1.ucVertical3DLine Line3d 
         Height          =   255
         Index           =   1
         Left            =   3780
         TabIndex        =   17
         Top             =   60
         Width           =   90
         _extentx        =   159
         _extenty        =   450
      End
      Begin Project1.ucVertical3DLine Line3d 
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   16
         Top             =   60
         Width           =   90
         _extentx        =   159
         _extenty        =   450
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   3
         Left            =   3900
         TabIndex        =   10
         Top             =   0
         Width           =   1035
         _extentx        =   1826
         _extenty        =   661
         caption         =   "&Justified"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":129C
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   0
         Left            =   180
         TabIndex        =   9
         ToolTipText     =   "Enviar por correio electrónico"
         Top             =   0
         Width           =   855
         _extentx        =   1508
         _extenty        =   661
         caption         =   "Send"
         picture         =   "Form1.frx":1836
         picturedisabled =   "Form1.frx":1DD0
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   1
         Left            =   1320
         TabIndex        =   8
         ToolTipText     =   "Esto es una prueba"
         Top             =   0
         Width           =   990
         _extentx        =   1746
         _extenty        =   661
         caption         =   "&Portrait"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":1F7A
      End
      Begin Project1.JCF_Button btn1 
         Height          =   375
         Index           =   2
         Left            =   2460
         TabIndex        =   7
         Top             =   0
         Width           =   1215
         _extentx        =   2143
         _extenty        =   661
         caption         =   "&Landscape"
         ischeckbutton   =   -1  'True
         picture         =   "Form1.frx":2514
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disabled"
      Height          =   390
      Index           =   4
      Left            =   4320
      TabIndex        =   4
      Top             =   1740
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OverDown"
      Height          =   390
      Index           =   3
      Left            =   3270
      TabIndex        =   3
      Top             =   1740
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Down"
      Height          =   390
      Index           =   2
      Left            =   2235
      TabIndex        =   2
      Top             =   1740
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Over"
      Height          =   390
      Index           =   1
      Left            =   1200
      TabIndex        =   1
      Top             =   1740
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Normal"
      Height          =   390
      Index           =   0
      Left            =   165
      TabIndex        =   0
      Top             =   1740
      Width           =   990
   End
   Begin VB.Label Label1 
      Caption         =   "Apply states on &send button:"
      Height          =   240
      Left            =   195
      TabIndex        =   5
      Top             =   1440
      Width           =   4275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NeoX As Single, IniWidth1 As Integer, IniWidth2 As Integer, selection As Integer

Private Sub Command1_Click(Index As Integer)
    btn1(0).State = Index
End Sub

Private Sub btn1_Click(Index As Integer)

    ' prevent reentrant calls (we have a DoEvents in here)
'    Static bBusy%
'    If bBusy Then Exit Sub
'    bBusy = True
    
    If Index = 1 Then
        If btn1(2).Value = 0 Then
            btn1(2).Value = 1
        Else
            btn1(2).Value = 0
        End If
    End If

    If Index = 2 Then
        If btn1(1).Value = 0 Then
            btn1(1).Value = 1
        Else
            btn1(1).Value = 0
        End If
    End If

'    bBusy = False
    'MsgBox "button " & Index & " pressed"
End Sub

Private Sub Form_Load()
    btn1(1).Value = 1
    IniWidth1 = Me.JCToolbar1(0).Width
    IniWidth2 = Me.JCToolbar1(1).Width
End Sub

Private Sub Form_Resize()
Dim i As Integer
    If (Me.WindowState = vbMinimized) Then
        Me.Visible = False
        bln_minimized = True
        Exit Sub
    Else    'not minimized
        MoveToolbar JCToolbar1(0), btn1, 0, 3, IniWidth1, Line3d, 0, 1
        MoveToolbar JCToolbar1(1), btn1, 4, 8, IniWidth2, Line3d, 2, 2
    End If
End Sub

Private Sub JCToolbar1_ButtonClick(Index As Integer)
    MsgBox "clicked"
End Sub

Private Sub JCToolbar1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    TimMove.Enabled = True
End If
selection = Index
End Sub

Private Sub JCToolbar1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    NeoX = JCToolbar1(Index).Left + X
    If NeoX < 0 Then
        NeoX = 0
    ElseIf NeoX > Me.Width - JCToolbar1(Index).Width Then
        NeoX = Me.Width - JCToolbar1(Index).Width - 100
    End If
End Sub

Private Sub JCToolbar1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    TimMove.Enabled = False
    
End Sub
Private Sub TimMove_Timer()
    JCToolbar1(selection).Left = NeoX
End Sub

Public Sub MoveToolbar(ToolBar As Object, BtnA As Object, BtnAFrom As Integer, BtnATo As Integer, IniWidth As Integer, Line3 As Object, Line3From As Integer, Line3To As Integer)
    Dim BlnHide As Boolean
    
    If Me.ScaleWidth < (ToolBar.Left + ToolBar.Width) Then
        If (Me.ScaleWidth - ToolBar.Left) > 450 Then
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
        For i = BtnAFrom To BtnATo
            If ToolBar.Width - 210 < BtnA(i).Left + BtnA(i).Width Then
                BtnA(i).Visible = False
                BlnHide = True
            Else
                BtnA(i).Visible = True
            End If
        Next i
        For i = Line3From To Line3To
            If ToolBar.Width - 210 < Line3(i).Left + Line3(i).Width Then
                Line3(i).Visible = False
            Else
                Line3(i).Visible = True
            End If
        Next i
    Else
        If (ToolBar.Left + IniWidth) < Me.ScaleWidth Then
            ToolBar.Width = IniWidth
        Else
            ToolBar.Width = Me.ScaleWidth - ToolBar.Left
        End If
        For i = BtnAFrom To BtnATo
            If ToolBar.Width - 210 < BtnA(i).Left + BtnA(i).Width Then
                BtnA(i).Visible = False
                BlnHide = True
            Else
                BtnA(i).Visible = True
            End If
        Next i
        For i = Line3From To Line3To
            If ToolBar.Width - 210 < Line3(i).Left + Line3(i).Width Then
                Line3(i).Visible = False
            Else
                Line3(i).Visible = True
            End If
        Next i
    End If
    ToolBar.IsStrech = BlnHide
End Sub
