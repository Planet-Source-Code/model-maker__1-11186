VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Editor - []"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   7515
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   7515
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   6015
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   2175
      Begin VB.CommandButton Command1 
         Caption         =   "Set name"
         Height          =   375
         Left            =   0
         TabIndex        =   14
         Top             =   3000
         Width           =   2055
      End
      Begin VB.Frame Frame1 
         Caption         =   "Stuff...?"
         Height          =   2415
         Left            =   120
         TabIndex        =   9
         Top             =   3480
         Width           =   1815
         Begin VB.OptionButton func 
            Caption         =   "Move Face"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   1335
         End
         Begin VB.OptionButton func 
            Caption         =   "Reposition"
            Height          =   255
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   1455
         End
         Begin VB.OptionButton func 
            Caption         =   "Select"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton func 
            Caption         =   "Move Vertex"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   13
            Top             =   840
            Width           =   1215
         End
         Begin VB.OptionButton func 
            Caption         =   "Add Box"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   12
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton func 
            Caption         =   "Move Box"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   11
            Top             =   1320
            Width           =   1335
         End
         Begin VB.OptionButton func 
            Caption         =   "Spin"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   10
            Top             =   1560
            Width           =   855
         End
      End
      Begin VB.ListBox Names 
         Height          =   2595
         Left            =   0
         TabIndex        =   8
         Tag             =   "0"
         Top             =   0
         Width           =   2055
      End
      Begin VB.TextBox NewName 
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   2640
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   0
      ScaleHeight     =   105
      ScaleWidth      =   105
      TabIndex        =   5
      Top             =   0
      Width           =   135
   End
   Begin VB.HScrollBar OfsHoris 
      Height          =   255
      LargeChange     =   200
      Left            =   0
      Max             =   2000
      Min             =   -2000
      SmallChange     =   20
      TabIndex        =   1
      Top             =   5760
      Value           =   -175
      Width           =   4935
   End
   Begin VB.VScrollBar ofVert 
      Height          =   5775
      LargeChange     =   200
      Left            =   4920
      Max             =   2000
      Min             =   -2000
      SmallChange     =   20
      TabIndex        =   2
      Top             =   0
      Value           =   -175
      Width           =   255
   End
   Begin VB.PictureBox RuleSide 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   0
      ScaleHeight     =   377
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   4
      Tag             =   "50"
      Top             =   120
      Width           =   135
   End
   Begin VB.PictureBox RuleTop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   329
      TabIndex        =   3
      Tag             =   "50"
      Top             =   0
      Width           =   4935
   End
   Begin MSComDlg.CommonDialog Load 
      Left            =   240
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdl"
   End
   Begin MSComDlg.CommonDialog Save 
      Left            =   840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.mdl"
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5655
      Left            =   120
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   317
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Line Guide 
         BorderStyle     =   2  'Dash
         Visible         =   0   'False
         X1              =   112
         X2              =   192
         Y1              =   160
         Y2              =   248
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOptions 
         Caption         =   "New"
         Index           =   1
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Load..."
         Index           =   2
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Merge with..."
         Index           =   3
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Save"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Save as..."
         Index           =   5
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Export..."
         Index           =   7
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuFileOptions 
         Caption         =   "Exit"
         Index           =   9
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditOptions 
         Caption         =   "Undo"
         Enabled         =   0   'False
         Index           =   1
         Shortcut        =   %{BKSP}
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "Paste"
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEditOptions 
         Caption         =   "Delete"
         Index           =   6
         Shortcut        =   {DEL}
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu oView 
         Caption         =   "View from top"
         Checked         =   -1  'True
         Index           =   1
         Shortcut        =   {F1}
      End
      Begin VB.Menu oView 
         Caption         =   "View from side"
         Index           =   2
         Shortcut        =   {F2}
      End
      Begin VB.Menu oView 
         Caption         =   "View from front"
         Index           =   3
         Shortcut        =   {F3}
      End
      Begin VB.Menu Space 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewTools 
         Caption         =   "Views toolbar"
         Index           =   1
         Shortcut        =   ^{F1}
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolOptions 
         Caption         =   "Set name"
         Index           =   1
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "Hide me"
         Index           =   2
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "Get info..."
         Index           =   3
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "Edit surface properties..."
         Index           =   4
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "3D viewer..."
         Index           =   5
      End
      Begin VB.Menu mnuToolOptions 
         Caption         =   "Open file viewer"
         Index           =   6
         Begin VB.Menu OpenViewer 
            Caption         =   "DirectX"
            Index           =   1
         End
         Begin VB.Menu OpenViewer 
            Caption         =   "Software (Texture mapping)"
            Index           =   2
         End
      End
   End
   Begin VB.Menu mnuPopup1 
      Caption         =   "PopUpMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup1Options 
         Caption         =   "Edit surface propeties"
         Index           =   1
      End
      Begin VB.Menu mnuPopup1Options 
         Caption         =   "Change colour"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MakeBox As Boolean
Dim startX, StartY
Dim bMoveX, bMoveY
Dim spinAngle As Integer
Dim StnColour As Double

Private Sub Command1_Click()
SetName
End Sub

Private Sub SetName()
Shape(Selected).Name = NewName.Text
Draw
End Sub


Private Sub Form_Load()
Zoom = 1
mnuViewTools(1).Checked = True
Views.Visible = True
Xofs = -OfsHoris.Value
Yofs = -ofVert.Value
Draw
RuleTop.Tag = Int(View.ScaleWidth / 2)
RuleSide.Tag = Int(View.ScaleHeight / 2)
Make_LookUp
DrawBars
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Views
End Sub

Private Sub func_Click(Index As Integer)
Draw
End Sub

Private Sub FillUndoBoard()

 UndoBoard.Colour = Shape(Selected).Colour
 
 For n = 1 To 8
  UndoBoard.Corners(n).X = Shape(Selected).Corners(n).X
  UndoBoard.Corners(n).Y = Shape(Selected).Corners(n).Y
  UndoBoard.Corners(n).z = Shape(Selected).Corners(n).z
 Next n
 
 For n = 1 To 6
  For m = 1 To 4
   UndoBoard.Face(n).Edge(m) = Shape(Selected).Face(n).Edge(m)
  Next m
 Next n
  UndoBoard.Name = Shape(Selected).Name
  UndoBoard.Used = True
  
  For n = 1 To 6
   UndoBoard.Face(n).Used = Shape(Selected).Face(n).Used
   UndoBoard.Face(n).Reverse = Shape(Selected).Face(n).Reverse
  Next n


End Sub

Private Sub mnuEditOptions_Click(Index As Integer)

Select Case Index

Case 1
  If UndoBoard.Used = False Then
   MsgBox "There is nothing to paste!", , "Wait..."
   Exit Sub
  End If

  Add = FindBox
  Shape(Add).Colour = UndoBoard.Colour
  For n = 1 To 8
   Shape(Add).Corners(n).X = UndoBoard.Corners(n).X - Ax
   Shape(Add).Corners(n).Y = UndoBoard.Corners(n).Y - Ay
   Shape(Add).Corners(n).z = UndoBoard.Corners(n).z - Az
  Next n
  For n = 1 To 6
   For m = 1 To 4
    Shape(Add).Face(n).Edge(m) = UndoBoard.Face(n).Edge(m)
   Next m
  Next n
  Shape(Add).Name = UndoBoard.Name
  Shape(Add).Used = True
 
  For n = 1 To 6
   Shape(Selected).Face(n).Used = UndoBoard.Face(n).Used
   Shape(Selected).Face(n).Reverse = UndoBoard.Face(n).Reverse
  Next n
  
  Shape(Selected).Used = False
  FillUndoBoard
  Selected = Add
  Draw
  mnuEditOptions(1).Caption = "Redo"


Case 3
 
 CopyBoard.Colour = Shape(Selected).Colour
 
 For n = 1 To 8
  CopyBoard.Corners(n).X = Shape(Selected).Corners(n).X
  CopyBoard.Corners(n).Y = Shape(Selected).Corners(n).Y
  CopyBoard.Corners(n).z = Shape(Selected).Corners(n).z
 Next n
 
 For n = 1 To 6
  For m = 1 To 4
   CopyBoard.Face(n).Edge(m) = Shape(Selected).Face(n).Edge(m)
  Next m
 Next n
  CopyBoard.Name = Shape(Selected).Name
  CopyBoard.Used = True
  
  For n = 1 To 6
   CopyBoard.Face(n).Used = Shape(Selected).Face(n).Used
   CopyBoard.Face(n).Reverse = Shape(Selected).Face(n).Reverse
  Next n
  
Case 4
 
 
 If CopyBoard.Used = False Then
  MsgBox "There is nothing to paste!", , "Wait..."
  Exit Sub
 End If
  
   Add = FindBox
  
If oView(1).Checked = True Then Ax = 10: Az = 10
If oView(2).Checked = True Then Ax = 10: Ay = 10
If oView(3).Checked = True Then Ay = 10: Az = 10
    
  Shape(Add).Colour = CopyBoard.Colour
  
 For n = 1 To 8
    Shape(Add).Corners(n).X = CopyBoard.Corners(n).X - Ax
    Shape(Add).Corners(n).Y = CopyBoard.Corners(n).Y - Ay
    Shape(Add).Corners(n).z = CopyBoard.Corners(n).z - Az
 Next n
 For n = 1 To 6
  For m = 1 To 4
   Shape(Add).Face(n).Edge(m) = CopyBoard.Face(n).Edge(m)
  Next m
 Next n
   Shape(Add).Name = CopyBoard.Name
   Shape(Add).Used = True
   
   Selected = Add
   
  For n = 1 To 6
   Shape(Selected).Face(n).Used = CopyBoard.Face(n).Used
   Shape(Selected).Face(n).Reverse = CopyBoard.Face(n).Reverse
  Next n
   
   
   Draw

Case 6
 
 
 FillUndoBoard
 mnuEditOptions(1).Caption = "Undo delete box"
 mnuEditOptions(1).Enabled = True
 Shape(Selected).Used = False
 Draw

  
End Select
End Sub

Private Sub mnuFileOptions_Click(Index As Integer)
Selected = 0
func(0).Value = True
Select Case Index
 Case 1 ' New!
  X = MsgBox("Are you sure you want to start a new model?", vbYesNo, "Wait...")
  If X = 7 Then Exit Sub
  ClearShapes
  Draw
  Load.Tag = ""
 Case 2 ' open
  Load.FileName = "*.mdl"
  Load.ShowOpen
  ClearShapes
  Call LoadNew(Load.FileName)
  Load.Tag = Load.FileName
 Case 3 ' merge - The same, but we don't clear memory first
  Load.ShowOpen
  
  Call LoadNew(Load.FileName)
  Load.Tag = Load.FileName
 Case 4 ' Save
   Call SaveAll(Load.Tag, 0)
   Form1.Tag = "Saved"
 Case 5 ' SaveAs
  Save.ShowSave
  If Mid(Save.FileName, Len(Save.FileName) - 3, 1) <> "." Then
   Save.FileName = Save.FileName + ".mdl"
  End If
  Call SaveAll(Save.FileName, 1)
  Load.Tag = Save.FileName
   Form1.Tag = "Saved"
 
 Case 7
  Form1.Visible = False
  Views.Visible = False
  ExportMe.Visible = True
  
  
 
 Case 9 ' Quit!
  If Form1.Tag = "Not Saved" Then
   message = "Your work has been changed since you last saved it. Do you want to save it before you exit?"
   X = MsgBox(message, vbYesNoCancel, "Wait...")
    
   If X = 7 Then End
   
   If X = 6 Then
    If mnuFileOptions(4).Enabled = True Then Call SaveAll(Load.Tag, 0): End
     
     Save.ShowSave
     Call SaveAll(Save.FileName, 1)
     Load.Tag = Save.FileName
     End
   End If
  Else
  X = MsgBox("Are you sure you want to quit?", vbYesNo, "Wait...")
  If X = 7 Then Exit Sub
  End
  End If
End Select

mnuFileOptions(4).Enabled = False
If Load.Tag <> "" Then mnuFileOptions(4).Enabled = True
Form1.Caption = "Editor - [" + Load.Tag + "]"
End Sub

Private Sub mnuPopup1Options_Click(Index As Integer)
Select Case Index
 Case 1
  ReFace.Visible = True
  ReFace.NowOn = Selected
  Names.Tag = Selected
  ReFace.start
 Case 2
  Load.ShowColor
  Shape(Selected).Colour = Load.Color
  StnColour = Load.Color
  Draw
End Select
End Sub

Private Sub mnuToolOptions_Click(Index As Integer)
Select Case Index
 Case 1
  SetName
 Case 2
  Views.Visible = False
  Form1.WindowState = 1
 Case 3
  GetData
 Case 4
  ReFace.Visible = True
  ReFace.start
 Case 5
  Viewer3D.Visible = True
  Form1.Visible = False
  Views.Visible = False
End Select
End Sub

Private Sub mnuViewTools_Click(Index As Integer)
If mnuViewTools(1).Checked = False Then
   mnuViewTools(1).Checked = True
   Views.Visible = True
  Else
   mnuViewTools(1).Checked = False
   Views.Visible = False
 End If
End Sub

Private Sub Names_click()
 If func(1) = True Then func(2).Value = True
 Selected = Nifty(Names.ListIndex + 1)
 NewName = Shape(Selected).Name
 Draw
End Sub

Private Sub OfsHoris_Change()
Xofs = -OfsHoris.Value
Draw
End Sub

Private Sub OfsHoris_Scroll()
Xofs = -OfsHoris.Value
Draw
End Sub

Private Sub ofVert_Change()
Yofs = -ofVert.Value
Draw
End Sub

Private Sub ofVert_Scroll()
Yofs = -ofVert.Value
Draw
End Sub

Private Sub OpenViewer_Click(Index As Integer)
Select Case Index
 Case 1
  s = App.Path + "\viewer.exe"
  Shell s, vbNormalFoc
 Case 2
  s = App.Path + "\TextureMapper.exe"
  Shell s, vbNormalFoc
End Select
End Sub



Private Sub oView_Click(Index As Integer)
For n = 1 To 3: oView(n).Checked = False: Next n
oView(Index).Checked = True
Draw
End Sub


Sub DrawBars()
RuleSide.Cls
RuleTop.Cls
For n = 2 - ((Yofs Mod 25) * 5.5) To RuleSide.ScaleHeight + ((Yofs Mod 25) * 4) Step 25
RuleSide.Line (6, n)-(8, n)
Next n

For n = 2 - ((Xofs Mod 25) * 5) To RuleTop.ScaleWidth + ((Xofs Mod 25) * 4) Step 25
RuleTop.Line (n, 6)-(n, 8)
Next n

Y = RuleSide.Tag
RuleSide.Line (0, Y - 3)-(0, Y + 3)
RuleSide.Line (1, Y - 2)-(1, Y + 2)
RuleSide.Line (2, Y - 1)-(2, Y + 1)



X = RuleTop.Tag
RuleTop.Line (X - 3, 0)-(X + 3, 0)
RuleTop.Line (X - 2, 1)-(X + 2, 1)
RuleTop.Line (X - 1, 2)-(X + 1, 2)
End Sub


Private Sub RuleSide_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RuleSide.Tag = Y
View.DrawStyle = 2
View.Line (0, Y)-(1000, Y)
View.DrawStyle = 0
DrawBars
End Sub

Private Sub RuleSide_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draw
End Sub

Private Sub RuleTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RuleTop.Tag = X
View.DrawStyle = 2
View.Line (X, 0)-(X, 1000)
View.DrawStyle = 0
DrawBars
End Sub

Private Sub RuleTop_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Draw
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If func(1) <> True And func(6) <> True Then
 FillUndoBoard
 If func(2) = True Then mnuEditOptions(1).Caption = "Undo move vertex"
 If func(3) = True Then mnuEditOptions(1).Caption = "Undo move face"
 If func(4) = True Then mnuEditOptions(1).Caption = "Undo move box"
 If func(5) = True Then mnuEditOptions(1).Caption = "Undo spin box"
 mnuEditOptions(1).Enabled = True
Else
 mnuEditOptions(1).Enabled = False
End If

Dim n As Byte, xx As Integer, zz As Integer

If Button = 2 Then
 mnuPopup1Options(1).Caption = "Edit surface properties for " + Shape(Selected).Name

 PopupMenu mnuPopup1
 Exit Sub
End If

If func(3) = True Then
  bMoveX = Xmouse
  bMoveY = Ymouse
 For n = 1 To 6
  CenX = 0: CenY = 0
  For m = 1 To 4
   
  If oView(1).Checked = True Then
   xx = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).X
   zz = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).z
  End If
 
 If oView(2).Checked = True Then
   xx = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).X
   zz = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).Y
 End If
 
 If oView(3).Checked = True Then
   xx = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).z
   zz = Shape(Selected).Corners(Shape(Selected).Face(n).Edge(m)).Y
 End If
   
  CenX = CenX + xx + Xofs
  CenY = CenY + zz + Yofs
 Next m
 CenX = CenX / 4
 CenY = CenY / 4
   
   If Xmouse > CenX - DisttoSelect Then
   If Xmouse < CenX + DisttoSelect Then
    If Ymouse > CenY - DisttoSelect Then
     If Ymouse < CenY + DisttoSelect Then
    
       Draw
       View.Circle (CenX, CenY), 4
       
       For s = 1 To 4
          Shape(Selected).Corners(Shape(Selected).Face(n).Edge(s)).MouseOver = True
       Next s
       If Shift = 1 Then Exit Sub
     End If
    End If
   End If
  End If
Next n
 
 
 Exit Sub
End If



'#######################################

If func(0) = True Then
  For m = 1 To 100
   For n = 1 To 8
   
  If oView(1).Checked = True Then
   xx = Shape(m).Corners(n).X
   zz = Shape(m).Corners(n).z
  End If
 
 If oView(2).Checked = True Then
   xx = Shape(m).Corners(n).X
   zz = Shape(m).Corners(n).Y
 End If
 
 If oView(3).Checked = True Then
   xx = Shape(m).Corners(n).z
   zz = Shape(m).Corners(n).Y
 End If
   
   xx = xx + Xofs
  zz = zz + Yofs
     
   If Xmouse > xx - DisttoSelect Then
   If Xmouse < xx + DisttoSelect Then
    If Ymouse > zz - DisttoSelect Then
     If Ymouse < zz + DisttoSelect Then
    
       'View.Circle (xx, zz), 4
       Selected = m
       Draw
       Exit Sub
     End If
    End If
   End If
  End If
  
  Next n
 Next m

End If




If func(5) = True Then

'##########

 For n = 1 To 8

 Shape(Selected).Corners(n).MouseOver = False
 
 If oView(1).Checked = True Then
   xx = Shape(Selected).Corners(n).X
   zz = Shape(Selected).Corners(n).z
 End If
 
 If oView(2).Checked = True Then
   xx = Shape(Selected).Corners(n).X
   zz = Shape(Selected).Corners(n).Y
 End If
 
 If oView(3).Checked = True Then
   xx = Shape(Selected).Corners(n).z
   zz = Shape(Selected).Corners(n).Y
 End If
 
  
  xx = xx + Xofs
  zz = zz + Yofs
  If Xmouse > xx - DisttoSelect Then
   If Xmouse < xx + DisttoSelect Then
    If Ymouse > zz - DisttoSelect Then
     If Ymouse < zz + DisttoSelect Then
       View.Circle (xx, zz), 4
       Shape(Selected).Corners(n).MouseOver = True
     End If
    End If
   End If
  End If
  
 Next n


'###########



 Guide.x1 = Xmouse
 Guide.Y1 = Ymouse
 Guide.Visible = True
 Exit Sub
End If

If func(4) = True Or func(6) = True Then
  bMoveX = Xmouse
  bMoveY = Ymouse
  Exit Sub
End If


If func(1) = True Then
 MakeBox = True
 startX = Xmouse
 StartY = Ymouse
 
 Guide.Visible = True
 Guide.x1 = Xmouse
 Guide.Y1 = Ymouse
 
End If


If func(2) = True Then
 Draw
 For n = 1 To 8

 Shape(Selected).Corners(n).MouseOver = False
 
 If oView(1).Checked = True Then
   xx = Shape(Selected).Corners(n).X
   zz = Shape(Selected).Corners(n).z
 End If
 
 If oView(2).Checked = True Then
   xx = Shape(Selected).Corners(n).X
   zz = Shape(Selected).Corners(n).Y
 End If
 
 If oView(3).Checked = True Then
   xx = Shape(Selected).Corners(n).z
   zz = Shape(Selected).Corners(n).Y
 End If
 
  
  xx = xx + Xofs
  zz = zz + Yofs
  If Xmouse > xx - DisttoSelect Then
   If Xmouse < xx + DisttoSelect Then
    If Ymouse > zz - DisttoSelect Then
     If Ymouse < zz + DisttoSelect Then
       View.Circle (xx, zz), 4
       Shape(Selected).Corners(n).MouseOver = True
       If Shift = 1 Then Exit Sub
   
     End If
    End If
   End If
  End If
  
 Next n
End If



End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)



Xmouse = (CInt(X / 5) * 5)
Ymouse = (CInt(Y / 5) * 5)
Guide.x2 = Xmouse
Guide.Y2 = Ymouse


If func(5) = True And Guide.Visible = True Then
 Guide.x2 = Xmouse
 Guide.Y2 = Ymouse
 'Guide.Visible = False
 
 TXX = Guide.x2 - Guide.x1
 Tyy = Guide.Y2 - Guide.Y1
 
 spinAngle = FindAngle(TXX, Tyy, 0)
 View.Line (Guide.x1, Guide.Y1)-(Guide.x1 + 28, Guide.Y1 + 12), RGB(255, 255, 0), BF
 View.Line (Guide.x1, Guide.Y1)-(Guide.x1 + 28, Guide.Y1 + 12), 0, B
 View.PSet (Guide.x1, Guide.Y1)
 View.Print spinAngle
 Exit Sub
End If


End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 2 Then
 
 Exit Sub
End If

'##############################################################

If func(5) = True Then
 If spinAngle < 0 Or spinAngle > 359 Then Exit Sub
 Guide.Visible = False
  
  
 For n = 1 To 8
  If Shape(Selected).Corners(n).MouseOver = True Then
   mx = Shape(Selected).Corners(n).X
   my = Shape(Selected).Corners(n).Y
   mz = Shape(Selected).Corners(n).z
   
  End If
 Next n
 
For n = 1 To 8
 mx = mx + Shape(Selected).Corners(n).X
 my = my + Shape(Selected).Corners(n).Y
 mz = mz + Shape(Selected).Corners(n).z
Next n

mx = mx / 8: my = my / 8: mz = mz / 8
 
 If oView(1).Checked = True Then Call Rotate(0, 360 - spinAngle, 0, Selected, mx, my, mz)
 If oView(3).Checked = True Then Call Rotate(360 - spinAngle, 0, 0, Selected, mx, my, mz)
 If oView(2).Checked = True Then Call Rotate(0, 0, 360 - spinAngle, Selected, mx, my, mz)
 
 
 For n = 1 To 8
  Shape(Selected).Corners(n).X = Rotated.Corners(n).X
  Shape(Selected).Corners(n).Y = Rotated.Corners(n).Y
  Shape(Selected).Corners(n).z = Rotated.Corners(n).z
 Next n
   

End If

'##############################################################

If func(6) = True Then
 Form1.Tag = "Not Saved"
 
 mx = Xmouse - bMoveX
 my = Ymouse - bMoveY
For m = 1 To 100
 For n = 1 To 8
  
  If oView(1).Checked = True Then
  Shape(m).Corners(n).X = Shape(m).Corners(n).X + mx
  Shape(m).Corners(n).z = Shape(m).Corners(n).z + my
  End If
 
 If oView(2).Checked = True Then
  Shape(m).Corners(n).X = Shape(m).Corners(n).X + mx
  Shape(m).Corners(n).Y = Shape(m).Corners(n).Y + my
 End If
   
 If oView(3).Checked = True Then
  Shape(m).Corners(n).z = Shape(m).Corners(n).z + mx
  Shape(m).Corners(n).Y = Shape(m).Corners(n).Y + my
 End If
  
 Next n
Next m
Draw
 Exit Sub


End If

If func(3) = True Then
 Form1.Tag = "Not Saved"
 mx = Xmouse - bMoveX
 my = Ymouse - bMoveY
 For n = 1 To 8
  If Shape(Selected).Corners(n).MouseOver = True Then
  If oView(1).Checked = True Then
   Shape(Selected).Corners(n).X = Shape(Selected).Corners(n).X + mx
   Shape(Selected).Corners(n).z = Shape(Selected).Corners(n).z + my
   End If
  If oView(2).Checked = True Then
   Shape(Selected).Corners(n).X = Shape(Selected).Corners(n).X + mx
   Shape(Selected).Corners(n).Y = Shape(Selected).Corners(n).Y + my
  End If
  If oView(3).Checked = True Then
   Shape(Selected).Corners(n).z = Shape(Selected).Corners(n).z + mx
   Shape(Selected).Corners(n).Y = Shape(Selected).Corners(n).Y + my
  End If
 End If
 Shape(Selected).Corners(n).MouseOver = False
 Next n
 Draw
 Exit Sub
End If



If func(4) = True Then
 Form1.Tag = "Not Saved"
 
 mx = Xmouse - bMoveX
 my = Ymouse - bMoveY
 For n = 1 To 8
  
 If oView(1).Checked = True Then
  Shape(Selected).Corners(n).X = Shape(Selected).Corners(n).X + mx
  Shape(Selected).Corners(n).z = Shape(Selected).Corners(n).z + my
  End If
 
 If oView(2).Checked = True Then
  Shape(Selected).Corners(n).X = Shape(Selected).Corners(n).X + mx
  Shape(Selected).Corners(n).Y = Shape(Selected).Corners(n).Y + my
 End If
   
 If oView(3).Checked = True Then
  Shape(Selected).Corners(n).z = Shape(Selected).Corners(n).z + mx
  Shape(Selected).Corners(n).Y = Shape(Selected).Corners(n).Y + my
 End If
  
 Next n

 Draw
 Exit Sub
End If


If func(1) = True And MakeBox = True Then
 Form1.Tag = "Not Saved"

 startX = startX - CInt((Xofs / 5)) * 5
 StartY = StartY - CInt((Yofs / 5)) * 5
 
 Xmouse = Xmouse - CInt((Xofs / 5)) * 5
 Ymouse = Ymouse - CInt((Yofs / 5)) * 5
 Guide.Visible = False
 Add = FindBox
 Shape(Add).Used = True
 Shape(Add).Colour = StnColour

If oView(1).Checked = True Then
 Shape(Add).Corners(1).X = startX: Shape(Add).Corners(1).Y = 50: Shape(Add).Corners(1).z = StartY
 Shape(Add).Corners(2).X = Xmouse: Shape(Add).Corners(2).Y = 50: Shape(Add).Corners(2).z = StartY
 Shape(Add).Corners(3).X = Xmouse: Shape(Add).Corners(3).Y = 50: Shape(Add).Corners(3).z = Ymouse
 Shape(Add).Corners(4).X = startX: Shape(Add).Corners(4).Y = 50: Shape(Add).Corners(4).z = Ymouse
 Shape(Add).Corners(5).X = startX: Shape(Add).Corners(5).Y = 150: Shape(Add).Corners(5).z = StartY
 Shape(Add).Corners(6).X = Xmouse: Shape(Add).Corners(6).Y = 150: Shape(Add).Corners(6).z = StartY
 Shape(Add).Corners(7).X = Xmouse: Shape(Add).Corners(7).Y = 150: Shape(Add).Corners(7).z = Ymouse
 Shape(Add).Corners(8).X = startX: Shape(Add).Corners(8).Y = 150: Shape(Add).Corners(8).z = Ymouse
End If

If oView(2).Checked = True Then
 Shape(Add).Corners(1).X = startX: Shape(Add).Corners(1).z = 150: Shape(Add).Corners(1).Y = StartY
 Shape(Add).Corners(2).X = Xmouse: Shape(Add).Corners(2).z = 150: Shape(Add).Corners(2).Y = StartY
 Shape(Add).Corners(3).X = Xmouse: Shape(Add).Corners(3).z = 150: Shape(Add).Corners(3).Y = Ymouse
 Shape(Add).Corners(4).X = startX: Shape(Add).Corners(4).z = 150: Shape(Add).Corners(4).Y = Ymouse
 Shape(Add).Corners(5).X = startX: Shape(Add).Corners(5).z = 50: Shape(Add).Corners(5).Y = StartY
 Shape(Add).Corners(6).X = Xmouse: Shape(Add).Corners(6).z = 50: Shape(Add).Corners(6).Y = StartY
 Shape(Add).Corners(7).X = Xmouse: Shape(Add).Corners(7).z = 50: Shape(Add).Corners(7).Y = Ymouse
 Shape(Add).Corners(8).X = startX: Shape(Add).Corners(8).z = 50: Shape(Add).Corners(8).Y = Ymouse
End If
 
If oView(3).Checked = True Then
 Shape(Add).Corners(1).z = startX: Shape(Add).Corners(1).X = 50: Shape(Add).Corners(1).Y = StartY
 Shape(Add).Corners(2).z = Xmouse: Shape(Add).Corners(2).X = 50: Shape(Add).Corners(2).Y = StartY
 Shape(Add).Corners(3).z = Xmouse: Shape(Add).Corners(3).X = 50: Shape(Add).Corners(3).Y = Ymouse
 Shape(Add).Corners(4).z = startX: Shape(Add).Corners(4).X = 50: Shape(Add).Corners(4).Y = Ymouse
 Shape(Add).Corners(5).z = startX: Shape(Add).Corners(5).X = 150: Shape(Add).Corners(5).Y = StartY
 Shape(Add).Corners(6).z = Xmouse: Shape(Add).Corners(6).X = 150: Shape(Add).Corners(6).Y = StartY
 Shape(Add).Corners(7).z = Xmouse: Shape(Add).Corners(7).X = 150: Shape(Add).Corners(7).Y = Ymouse
 Shape(Add).Corners(8).z = startX: Shape(Add).Corners(8).X = 150: Shape(Add).Corners(8).Y = Ymouse
End If
 
 
 For n = 1 To 6
  Shape(Add).Face(n).Used = True
  Shape(Add).Face(n).Reverse = False
 Next n
 
 
 Shape(Add).Name = "Cube " & Add
 Selected = Add
 Call FillCube(Add)
End If
MakeBox = False

If func(2) = True Then
 Form1.Tag = "Not Saved"

 Xmouse = Xmouse - CInt((Xofs / 5)) * 5
 Ymouse = Ymouse - CInt((Yofs / 5)) * 5
 
  For n = 1 To 8
  If Shape(Selected).Corners(n).MouseOver = True Then
  
 Shape(Selected).Corners(n).MouseOver = False
 
 If oView(1).Checked = True Then
   Shape(Selected).Corners(n).X = Xmouse
   Shape(Selected).Corners(n).z = Ymouse
 End If
 
 If oView(2).Checked = True Then
   Shape(Selected).Corners(n).X = Xmouse
   Shape(Selected).Corners(n).Y = Ymouse
 End If
   
 If oView(3).Checked = True Then
   Shape(Selected).Corners(n).z = Xmouse
   Shape(Selected).Corners(n).Y = Ymouse
 End If
   

  End If
 Next n
End If

Draw
End Sub

Sub Draw()
View.Cls
NewName.Text = Shape(Selected).Name
DrawBars
View.Line (0, Yofs)-(1000, Yofs)
View.Line (Xofs, 0)-(Xofs, 1000)


If oView(1).Checked = True Then DrawFromTop
If oView(2).Checked = True Then DrawFromSide
If oView(3).Checked = True Then DrawFromFront
End Sub
