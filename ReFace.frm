VERSION 5.00
Begin VB.Form ReFace 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Surface Editor"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7125
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5160
   ScaleWidth      =   7125
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Height          =   4935
      Left            =   4200
      TabIndex        =   15
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Skip to >>"
      Height          =   375
      Left            =   2880
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remove hidden faces"
      Height          =   375
      Left            =   5520
      TabIndex        =   9
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Reverse visibility"
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   4560
      Width           =   1335
   End
   Begin VB.ListBox ReNorm 
      Height          =   1410
      Left            =   5520
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.PictureBox SB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   120
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   4
      Top             =   720
      Width           =   5175
   End
   Begin VB.ListBox Faces 
      Height          =   1410
      ItemData        =   "ReFace.frx":0000
      Left            =   5520
      List            =   "ReFace.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Make all visible"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "Next  >"
      Height          =   375
      Index           =   1
      Left            =   1440
      TabIndex        =   1
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdMove 
      Caption         =   "<  Back"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label NowOn 
      Alignment       =   2  'Center
      Caption         =   "99"
      Height          =   255
      Left            =   1800
      TabIndex        =   13
      Top             =   360
      Width           =   255
   End
   Begin VB.Label BName 
      Caption         =   "Hi there, punk"
      Height          =   255
      Left            =   2160
      TabIndex        =   12
      Top             =   360
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Currently working on:"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label NumToDo 
      Caption         =   "No-"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label4 
      Caption         =   "Reverse visiblity:"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   2160
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Visible faces:"
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "ReFace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Xmouse, Ymouse As Integer
Dim Xof, Yof As Integer
Dim Sca As Single
Dim Angle1, Angle2, Angle3 As Integer

Sub start()

For n = 1 To 6
 Faces.AddItem "Face " & n
 ReNorm.AddItem "Face " & n
Next n
 


 If CountBox = 0 Then
  MsgBox "You have to have at least one brush in your model before using this function", , "Wait.."
   Form1.Enabled = True
   If Views.Visible = True Then Views.Enabled = True
   Unload Me
   Exit Sub
 End If
 NumToDo = "Number of bushes to do : " + Str(CountBox)
 
 NowOn = 1
 
 If Form1.Names.Tag <> 0 Then
  NowOn = Form1.Names.Tag
 
 End If
 Form1.Names.Tag = 0
 

 If Views.Visible = True Then Views.Enabled = False
 BName = Shape(Nifty(NowOn)).Name
 Selected = Nifty(NowOn)
 Form1.Draw
 
 If CountBox = 1 Then cmdMove(1).Enabled = False

 Form1.View.Enabled = False
 Form1.mnuEdit.Enabled = False
 Form1.mnuFile.Enabled = False
 Form1.mnuTools.Enabled = False
 Form1.mnuView.Enabled = False
 Yof = SB.ScaleHeight / 2
 Xof = SB.ScaleWidth / 2
 Sca = 1
 
 For m = 1 To 6
 Faces.Selected(m - 1) = False
 If Shape(NowOn).Face(m).Used = True Then
   Faces.Selected(m - 1) = True
 End If
Next m

For m = 1 To 6
 ReNorm.Selected(m - 1) = False
 If Shape(NowOn).Face(m).Reverse = True Then
   ReNorm.Selected(m - 1) = True
 End If
Next m
 
 If NowOn = CountBox Then cmdMove(1).Enabled = False
 If NowOn = 1 Then cmdMove(0).Enabled = False
 
 
 DrawBrush
End Sub

Private Sub Check1_Click()
DrawBrush
End Sub

Private Sub cmdMove_Click(Index As Integer)
 If Index = 0 Then NowOn = NowOn - 1
 If Index = 1 Then NowOn = NowOn + 1



 BName = Shape(Nifty(NowOn)).Name
 Selected = Nifty(NowOn.Caption)
 Form1.Draw
 DrawBrush
 
 cmdMove(0).Enabled = True
 cmdMove(1).Enabled = True
 
 
 If Index = 1 And NowOn = CountBox Then cmdMove(1).Enabled = False
 If Index = 0 And NowOn = 1 Then cmdMove(0).Enabled = False
 
 For m = 1 To 6
 Faces.Selected(m - 1) = False
 If Shape(NowOn).Face(m).Used = True Then
   Faces.Selected(m - 1) = True
 End If
Next m

For m = 1 To 6
 ReNorm.Selected(m - 1) = False
 If Shape(NowOn).Face(m).Reverse = True Then
   ReNorm.Selected(m - 1) = True
 End If
Next m
 
End Sub


Private Sub Command1_Click()
X = MsgBox("This will make all faces visible. Are you sure you want to continue?", vbYesNo + 48, "Wait")
If X = 7 Then Exit Sub
For n = 1 To 100
 If Shape(n).Used = True Then
   For m = 1 To 6
   Shape(n).Face(m).Used = True
   Faces.Selected(m - 1) = True
  Next m
 End If
Next n

DrawBrush
End Sub

Private Sub Command2_Click()
For n = 0 To 5
 If ReNorm.Selected(n) = True Then
   ReNorm.Selected(n) = False
   Shape(NowOn).Face(n).Reverse = False
  Else
   Shape(NowOn).Face(n).Reverse = True
   ReNorm.Selected(n) = True
  End If
Next n
SB.Cls
DrawBrush
End Sub

Private Sub Command3_Click()
List1.Visible = True
List1.Clear
For n = 1 To 100
 If Shape(n).Used = True Then
  List1.AddItem Shape(n).Name
 End If
Next n
List1.SetFocus
End Sub

Private Sub Faces_Click()
Shape(NowOn).Face(Faces.ListIndex + 1).Used = Faces.Selected(Faces.ListIndex)
DrawBrush
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Views.Visible = True Then Views.Enabled = True
 Form1.View.Enabled = True
 Form1.mnuView.Enabled = True
 Form1.mnuEdit.Enabled = True
 Form1.mnuFile.Enabled = True
 Form1.mnuTools.Enabled = True
End Sub

Sub DrawBrush()

For n = 1 To 8
 mx = mx + Shape(NowOn).Corners(n).X
 my = my + Shape(NowOn).Corners(n).Y
 mz = mz + Shape(NowOn).Corners(n).z
Next n

mx = mx / 8: my = my / 8: mz = mz / 8

Zeye = 800

Dim ner(4, 2) As Integer
FDraw = False

Call Rotate(Angle1, Angle2, Angle3, NowOn, mx, my, mz)

SB.Cls
For m = 1 To 6
 For n = 1 To 4
  X = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).X
  Y = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).Y
  z = Rotated.Corners(Shape(NowOn).Face(m).Edge(n)).z
  ner(n, 1) = Xof + Int(X * (Zeye / (Zeye - z))) * Sca
  ner(n, 2) = Yof + Int(Y * (Zeye / (Zeye - z))) * Sca
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
     
     DrawThisFace = False
     If NormalCheck = 1 And Normal > 0 Then DrawThisFace = True
     If NormalCheck = -1 And Normal < 0 Then DrawThisFace = True
     If Check1.Value = 0 Then DrawThisFace = True
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
        SB.Line (x1, Y1)-(x2, Y2)
        
          If FDraw = False Then
              FDraw = True
               High = Y1: low = Y1
               fleft = x1: fright = x1
          End If
        
    End If
  
  

  
  If Y1 > low Then low = Y1
  If Y2 > low Then low = Y2
  If Y1 < High Then High = Y1
  If Y2 < High Then High = Y2
  
  If x1 > fleft Then fleft = x1
  If x2 > fleft Then fleft = x2
  If x1 < fright Then fright = x1
  If x2 < fright Then fright = x2
 
  
  
  Next n
  If DrawThisFace = True Then
  Xmid = Xmid / 4
  Ymid = Ymid / 4
  SB.ForeColor = vbRed
  SB.PSet (Xmid, Ymid)
  SB.Print m
  SB.PSet (Xmid, Ymid + 1)
  SB.PSet (Xmid - 1, Ymid)
  SB.PSet (Xmid + 1, Ymid)
  SB.PSet (Xmid, Ymid - 1)
 End If
   
  SB.ForeColor = vbBlack
Next m

Xmid = (fleft + fright) / 2
Ymid = (High + low) / 2

   SB.Line (Xmid, Ymid)-(SB.ScaleWidth / 2, SB.ScaleHeight / 2), RGB(200, 200, 200)

End Sub



Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub List1_Click()
NowOn = Nifty(List1.ListIndex + 1)
List1.Visible = False
DrawBrush
cmdMove(0).Enabled = True
If NowOn = 1 Then cmdMove(0).Enabled = False
cmdMove(1).Enabled = True
If NowOn = CountBox Then cmdMove(1).Enabled = False

End Sub

Private Sub List1_LostFocus()
List1.Visible = False
End Sub

Private Sub ReNorm_Click()
Shape(NowOn).Face(ReNorm.ListIndex + 1).Reverse = ReNorm.Selected(ReNorm.ListIndex)
DrawBrush
End Sub

Private Sub SB_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 And Shift = 0 Then
 Angle1 = Angle1 + (Ymouse - Y)
 Angle2 = Angle2 + (Xmouse - X)
 Angle1 = Angle1 Mod 359
 Angle2 = Angle2 Mod 359
 If Angle1 < 0 Then Angle1 = Angle1 + 360
 If Angle2 < 0 Then Angle2 = Angle2 + 360
 
 DrawBrush
End If


If Button = 1 And Shift = 0 Then
 Xof = Xof - (Xmouse - X)
 Yof = Yof - (Ymouse - Y)
 DrawBrush
End If
If Button = 1 And Shift = 1 Then
 Sca = Sca + (Xmouse - X) / 40
 Sca = Abs(Sca)
 DrawBrush
End If

Xmouse = X
Ymouse = Y

End Sub
