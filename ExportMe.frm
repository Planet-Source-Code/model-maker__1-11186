VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form ExportMe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar TimeLeft 
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3360
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog Export 
      Left            =   3480
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton runExport 
      Caption         =   "Export..."
      Height          =   375
      Left            =   5400
      TabIndex        =   6
      Top             =   2640
      Width           =   1815
   End
   Begin VB.CommandButton GetFile 
      Caption         =   "Select file"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame2 
      Caption         =   "Format"
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   7095
      Begin VB.CheckBox removeTemp 
         Caption         =   "Remove tempory files?"
         Height          =   255
         Left            =   4320
         TabIndex        =   9
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.OptionButton Advanced 
         Caption         =   "Direct X mode - Corners and faces seperate"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Value           =   -1  'True
         Width           =   3495
      End
      Begin VB.OptionButton Normal 
         Caption         =   "Simple mode - Seperate faces"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label FileName 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "c:\DefaultModel.dat"
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   480
      Width           =   6975
   End
   Begin VB.Label Label1 
      Caption         =   "Export to :-"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "ExportMe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
ExtraInfo.Enabled = False
If Check1.Value = 1 Then ExtraInfo.Enabled = True
End Sub

Private Sub GetFile_Click()
 Export.FileName = "*.dat"
 Export.ShowSave
 If Mid(Export.FileName, Len(Export.FileName) - 3, 1) <> "." Then
  Export.FileName = Export.FileName + ".dat"
 End If
 FileName = Export.FileName
End Sub

Private Sub Cancel_Click()
 Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Form1.Visible = True
 If Form1.mnuViewTools(1).Checked = True Then Views.Visible = True
 Form1.SetFocus
End Sub

Private Sub runExport_Click()
' On Error GoTo ErrorNo1
 Open FileName For Output As #1
' On Error GoTo ErrorNo2
 
 
 BasicExport
 
 If Advanced = True Then
  ConvertFormat
 Else
 
    msage = ""
    msage = msage + "Basic Compiler finished..." + vbNewLine + vbNewLine
    msage = msage + "Model info..." + vbNewLine + vbNewLine
    msage = msage + "Total number of corners    :   " & CountFace * 4 & vbNewLine
    msage = msage + "Total number of faces       :   " & CountFace & vbNewLine & vbNewLine
    MsgBox msage
    
 End If
 
 Close

Exit Sub
ErrorNo1: MsgBox "Couldn't save to that file."
Exit Sub
ErrorNo2: MsgBox "Unexpected error on "
End Sub
