VERSION 5.00
Begin VB.Form Viewer3D 
   ClientHeight    =   4320
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   381
      TabIndex        =   0
      Top             =   0
      Width           =   5775
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Begin VB.Menu mnuOptionsOptions 
         Caption         =   "Dimensions"
         Index           =   1
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOptionsOptions 
         Caption         =   "Remove hidden faces"
         Index           =   2
      End
      Begin VB.Menu mnuOptionsOptions 
         Caption         =   "Use Colours"
         Index           =   3
      End
      Begin VB.Menu mnuOptionsOptions 
         Caption         =   "Exit"
         Index           =   4
      End
   End
End
Attribute VB_Name = "Viewer3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NowOn As Byte
Dim Xof As Integer, Yof As Integer
Dim Xmouse As Integer, Ymouse As Integer
Dim Xstart As Integer, Ystart As Integer
Dim Angle1, Angle2, Angle3


Private Sub Form_Load()
Me.Caption = "3D Viewer - [" + Form1.Load.Tag + "]"
End Sub

Private Sub Form_Resize()
View.Width = Me.ScaleWidth
View.Height = Me.ScaleHeight
Xof = View.ScaleWidth / 2
Yof = View.ScaleHeight / 2
DrawBrush
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form1.Visible = True
 If Form1.mnuViewTools(1).Checked = True Then
  Views.Visible = True
 End If
 Form1.SetFocus
End Sub


Sub DrawBrush()

View.Cls

Zeye = 800

Dim ner(4, 2) As Integer
FDraw = False

 For NowOn = 1 To 100
  If Shape(NowOn).Used = True Then
   GoSub DrawThisBrush
  End If
 Next NowOn
Exit Sub


DrawThisBrush:




Call Rotate(Angle1, Angle2, Angle3, NowOn, mx, my, mz)


For m = 1 To 6
 For n = 1 To 4
  X = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).X
  Y = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).Y
  z = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).z
  ner(n, 1) = Xof + Int(X * (Zeye / (Zeye - z))) * 1
  ner(n, 2) = Yof + Int(Y * (Zeye / (Zeye - z))) * 1
 Next n
 Xmid = 0: Ymid = 0
  For n = 1 To 4
  
  
If Shape(NowOn).Face(m).Reverse = True Then NormalCheck = 1
If Shape(NowOn).Face(m).Reverse = False Then NormalCheck = -1
  
     xx1 = ner(1, 1)
     xx2 = ner(2, 1)
     xx3 = ner(3, 1)
     yy1 = ner(1, 2)
     yy2 = ner(2, 2)
     yy3 = ner(3, 2)
     Normal = ((yy1 - yy3) * (xx2 - xx1) - (xx1 - xx3) * (yy2 - yy1))
     
     DrawThisFace = True
     If NormalCheck = 1 And Normal < 0 Then DrawThisFace = False
     If NormalCheck = -1 And Normal > 0 Then DrawThisFace = False
     If mnuOptionsOptions(2).Checked = False Then DrawThisFace = True
     
     If Shape(NowOn).Face(m).Used = False Then DrawThisFace = False
     
     If DrawThisFace = True Then
        If n = 4 Then
            x1 = ner(n, 1)
            Y1 = ner(n, 2)
            x2 = ner(1, 1)
            Y2 = ner(1, 2)
        Else
            x1 = ner(n, 1)
            Y1 = ner(n, 2)
            x2 = ner(n + 1, 1)
            Y2 = ner(n + 1, 2)
        End If
        Xmid = Xmid + x1
        Ymid = Ymid + Y1
        If mnuOptionsOptions(3).Checked = True Then
         View.ForeColor = Shape(NowOn).Colour
        End If
        View.Line (x1, Y1)-(x2, Y2)
    End If
  Next n
Next m

Return

End Sub


Private Sub mnuOptionsOptions_Click(Index As Integer)
Select Case Index

 Case 1, 2, 3
  If mnuOptionsOptions(Index).Checked = True Then
    mnuOptionsOptions(Index).Checked = False
   Else
    mnuOptionsOptions(Index).Checked = True
   End If
   DrawBrush
 Case 4
  Unload Me

End Select
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xstart = Xmouse
Ystart = Ymouse
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Xmouse = X
Ymouse = Y
If Button = 1 Then
End If

End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 1 Then
 Xof = Xof + (Xmouse - Xstart)
 Yof = Yof + (Ymouse - Ystart)
 DrawBrush
End If

If Button = 2 Then
 Angle1 = Angle1 + (Ystart - Ymouse)
 Angle2 = Angle2 + (Xstart - Xmouse)
 Angle1 = Angle1 Mod 359
 Angle2 = Angle2 Mod 359
 If Angle1 < 0 Then Angle1 = Angle1 + 360
 If Angle2 < 0 Then Angle2 = Angle2 + 360
 DrawBrush
End If
End Sub

