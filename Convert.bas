Attribute VB_Name = "Convert"
Dim tima As Integer
Dim X As Integer, Y As Integer, z As Integer

Sub ConvertFormat()
tima = 0
Close

rawfile = ExportMe.FileName

Open rawfile For Input As #1
Open "output.txt" For Output As #2
Points = 0
Input #1, Totalfaces
Totalfaces = Totalfaces - 1
For m = 1 To Totalfaces
 Input #1, Edge ', Red, Green, Blue
 totedge = totedge + Edge
 For n = 1 To Edge
   Input #1, X, Y, z
   Print #2, X,
   Print #2, Y,
   Print #2, z
   Points = Points + 1
  Next n
Next m
Close
                   
laps = totedge + totedge + Totalfaces + 1


Open "output.txt" For Input As #1
Open "output2.txt" For Output As #2
For n = 1 To Points
 Input #1, X, Y, z
 Print #2, X, Y, z
 Update (laps)
Next n
Close

Open "output.txt" For Input As #1
Open "output3.txt" For Output As #3

For n = 1 To Points
 Update (laps)
Found = 0
Input #1, X, Y, z


Open "output2.txt" For Input As #2
For m = 1 To n - 1
 Input #2, xx, yy, zz
 If X = xx And Y = yy And z = zz Then
  Found = 1
 End If
Next m
 
Close #2
 
If Found = 0 Then
 Print #3, X, Y, z
 NewPoints = NewPoints + 1
End If
Next n
Close

Open rawfile For Input As #1
Open "output4.txt" For Output As #2
Input #1, Totalfaces
For m = 1 To Totalfaces
 Input #1, Edge ', Red, Green, Blue
  Print #2, Edge,
  For n = 1 To Edge
   Input #1, X, Y, z
   X = Int(X): Y = Int(Y): z = Int(z)
    Open "output3.txt" For Input As #3
     
     For ifnd = 0 To NewPoints - 1
      Input #3, xx, yy, zz
      xx = Int(xx): yy = Int(yy): zz = Int(zz)
      If X = xx And Y = yy And z = zz Then
       Print #2, ifnd,
      End If
     Next ifnd
      
    Close #3
  Next n
  Print #2,
Next m
Close

rawfile = ExportMe.FileName


Open "output3.txt" For Input As #1
Open "output4.txt" For Input As #2
Open rawfile For Output As #3

'Open "text.txt" For Output As #4

Print #3, NewPoints - 1
Print #3, (Totalfaces * 5)

For n = 1 To NewPoints
 Input #1, X, Y, z
 X = Int(X): Y = Int(Y): z = Int(z)
 Print #3, X,
 Print #3, z,
 Print #3, Y
Next n

For n = 1 To Totalfaces
Line Input #2, xx
Print #3, xx
   Update (laps)
Next n

Close

If ExportMe.removeTemp = 1 Then
 Kill "output.txt"
 Kill "output2.txt"
 Kill "output3.txt"
 Kill "output4.txt"
End If

msage = ""
msage = msage + "Advanced Compiler finished..." + vbNewLine + vbNewLine
msage = msage + "Model info..." + vbNewLine + vbNewLine
msage = msage + "Total number of corners    :   " & NewPoints & vbNewLine
msage = msage + "Total number of edges      :   " & Totalfaces & vbNewLine & vbNewLine
msage = msage + "Do you want to open the DirectX viewer now?"

X = MsgBox(msage, vbYesNo)
s = App.Path + "\viewer"
If X = vbYes Then Shell s, vbNormalFocus
End Sub


Sub Update(laps)
 tima = tima + 1
 num = Int((tima / laps) * 100)
 ExportMe.TimeLeft.Value = num
End Sub


