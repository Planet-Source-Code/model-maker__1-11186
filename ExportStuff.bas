Attribute VB_Name = "ExportStuff"
Public Sub GetData()
 tb = vbTab
 message = message + "File name" + tb + ":" + tb + Form1.Load.Tag + vbNewLine
 message = message + Space(60) + vbNewLine
 message = message + "Brushes       " + tb + ":" + tb + Str(Form1.Names.ListCount) + vbNewLine
 message = message + "Estimated Brushes" + tb + ":" + tb + Str(Form1.Names.ListCount * 6) + vbNewLine
 message = message + "Estimated vertices" + tb + ":" + tb + Str(Form1.Names.ListCount * 8) + vbNewLine
 message = message + vbNewLine
 message = message + vbNewLine
 X = MsgBox(message, 48, "Model info...")
End Sub


Sub BasicExport()
Print #1, CountFace

For n = 1 To 100
 If Shape(n).Used = True Then

For m = 1 To 6
   
If Shape(n).Face(m).Used = True Then
Print #1, 4;
 If Shape(n).Face(m).Reverse = True Then
  For s = 1 To 4
   X = Shape(n).Corners(Shape(n).Face(m).Edge(s)).X
   Y = Shape(n).Corners(Shape(n).Face(m).Edge(s)).Y
   z = Shape(n).Corners(Shape(n).Face(m).Edge(s)).z
   Print #1, X; " "; Y; " "; z,
  Next s
 Else
   For s = 4 To 1 Step -1
   X = Shape(n).Corners(Shape(n).Face(m).Edge(s)).X
   Y = Shape(n).Corners(Shape(n).Face(m).Edge(s)).Y
   z = Shape(n).Corners(Shape(n).Face(m).Edge(s)).z
   Print #1, X; " "; Y; " "; z,
  Next s
 End If
 Print #1,
End If

 
Next m


 End If
Next n


End Sub



