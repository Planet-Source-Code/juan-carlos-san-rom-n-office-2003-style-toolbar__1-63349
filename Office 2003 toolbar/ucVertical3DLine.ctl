VERSION 5.00
Begin VB.UserControl ucVertical3DLine 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BackStyle       =   0  'Transparent
   ClientHeight    =   570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ScaleHeight     =   570
   ScaleWidth      =   240
   ToolboxBitmap   =   "ucVertical3DLine.ctx":0000
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   15
      X2              =   15
      Y1              =   15
      Y2              =   510
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   570
   End
End
Attribute VB_Name = "ucVertical3DLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Sub UserControl_Resize()
    UserControl.Width = 90
    'UserControl.Height = 250
    Line1(0).BorderColor = &H808080      'ColorToolbar
    Line1(0).Y2 = UserControl.Height
    Line1(1).Y2 = UserControl.Height
End Sub
