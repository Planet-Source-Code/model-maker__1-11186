VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cube"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6045
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog FileName 
      Left            =   840
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox View 
      Height          =   4335
      Left            =   0
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   397
      TabIndex        =   5
      Top             =   -120
      Width           =   6015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Press me to use textures"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   2655
   End
   Begin VB.CheckBox opt2 
      Caption         =   "Texture mapped"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar A 
      Height          =   255
      Index           =   3
      LargeChange     =   22
      Left            =   3240
      Max             =   0
      Min             =   359
      TabIndex        =   2
      Top             =   5040
      Width           =   2655
   End
   Begin VB.HScrollBar A 
      Height          =   255
      Index           =   2
      LargeChange     =   22
      Left            =   3240
      Max             =   0
      Min             =   359
      TabIndex        =   1
      Top             =   4680
      Width           =   2655
   End
   Begin VB.HScrollBar A 
      Height          =   255
      Index           =   1
      LargeChange     =   22
      Left            =   3240
      Max             =   0
      Min             =   359
      TabIndex        =   0
      Top             =   4320
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub A_Change(Index As Integer)
 opt2 = 0
 View.Cls
 Entity.Angle.X = A(1)
 Entity.Angle.Y = A(2)
 Entity.Angle.Z = A(3)
 Entity.Model = 1
 DrawModel
End Sub

Private Sub A_Scroll(Index As Integer)
 View.Cls
 Entity.Angle.X = A(1)
 Entity.Angle.Y = A(2)
 Entity.Angle.Z = A(3)
 Entity.Model = 1
 opt2 = 0
 DrawModel
End Sub

Private Sub Command1_Click()
 opt2 = 1
 View.Cls
 Entity.Angle.X = A(1)
 Entity.Angle.Y = A(2)
 Entity.Angle.Z = A(3)
 Entity.Model = 1
 DrawModel
End Sub

Private Sub Form_Load()
 FileName.FileName = "*.dat"
 FileName.ShowOpen
 Setup
 Make_LookUp
 GetTexture
 Random_World
End Sub

