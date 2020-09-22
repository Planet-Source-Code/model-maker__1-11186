VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Maker - Viewer"
   ClientHeight    =   5610
   ClientLeft      =   5955
   ClientTop       =   5490
   ClientWidth     =   5535
   FillColor       =   &H00404040&
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar hScale 
      Height          =   255
      Left            =   1080
      Max             =   100
      TabIndex        =   4
      Top             =   5160
      Value           =   10
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton Load 
      Caption         =   "Load"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   4680
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdLoad 
      Left            =   2520
      Top             =   5040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      FileName        =   "*.dat"
      Filter          =   "raw"
   End
   Begin VB.CommandButton BUTN_Quit 
      Caption         =   "Quit"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   4680
      Width           =   1335
   End
   Begin VB.PictureBox View 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   4575
      Left            =   0
      ScaleHeight     =   301
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   365
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MD3D As Boolean         'holder for mouse down in 3d window

Dim XStart As Integer       'position user first clicked on
Dim YStart As Integer

Dim VertArray() As D3DVECTOR 'vertices for the object
Dim SideFaces() As Long     'vertices making the sides of each face
        
Private Sub BUTN_Quit_Click()
    End
End Sub

Private Sub Command1_Click()
    InitRM View
    InitScene
    RenderScene
End Sub

Private Sub Form_Load()

    XStart = -1
    YStart = -1
    
    'show the form so directx has a surface to draw to
    Show
    
    'do events to handle mousemovement and other system tasks
    DoEvents
    
    'initialize directx in retained mode (pass in the pict box
    'that will contain the 3d object
    InitRM View
    
    'initialize the scene
    InitScene
    
    'draw the scene
    RenderScene
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'cleanup the directdraw variabes and memory allocations
    CleanUp
End Sub

Private Sub Load_Click()
cdLoad.FileName = "*.dat"
cdLoad.ShowOpen
cLoadModels
End Sub

Private Sub View_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MD3D = True
    m_LastX = X
    m_LastY = Y
    
End Sub

Private Sub View_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MD3D = True Then
        RotateTrackBall CInt(X), CInt(Y)
        RenderScene
    End If
End Sub

Private Sub View_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MD3D = False
End Sub


Private Sub cLoadModels()
On Error GoTo Knackered

Open cdLoad.FileName For Input As #1
On Error GoTo knackered2
Input #1, TotalPoints
Input #1, TotalEdges
ReDim VertArray(TotalPoints)
ReDim SideFaces(TotalEdges)

For Tp = 0 To TotalPoints
  Input #1, VertArray(Tp).X
  Input #1, VertArray(Tp).Y
  Input #1, VertArray(Tp).z
  VertArray(Tp).X = VertArray(Tp).X * (hScale / 100)
  VertArray(Tp).Y = VertArray(Tp).Y * (hScale / 100)
  VertArray(Tp).z = VertArray(Tp).z * (hScale / 100)
  
  
Next Tp

For Te = 0 To TotalEdges - 1
 Input #1, SideFaces(Te)
Next Te


Close

AddShape TotalPoints + 1, VertArray, SideFaces

RenderScene
Exit Sub
Knackered:

Exit Sub
knackered2:
MsgBox "Agg, an error - Your model don't work right!"

End Sub

