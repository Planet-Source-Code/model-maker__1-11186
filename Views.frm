VERSION 5.00
Begin VB.Form Views 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Views Toolbar"
   ClientHeight    =   585
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5175
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   585
   ScaleWidth      =   5175
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton V 
      Caption         =   "View from Front"
      Height          =   375
      Index           =   2
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton V 
      Caption         =   "View from Side"
      Height          =   375
      Index           =   1
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.OptionButton V 
      Caption         =   "View from Top"
      Height          =   375
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Views"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Unload(Cancel As Integer)
Form1.mnuViewTools(1).Checked = False
End Sub

Private Sub V_Click(Index As Integer)
For n = 1 To 3
 Form1.oView(n).Checked = False
Next n
Form1.oView(Index + 1).Checked = True
Form1.Draw
Form1.SetFocus
End Sub
