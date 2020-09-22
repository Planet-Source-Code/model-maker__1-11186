Attribute VB_Name = "Load_N_Save"
Public Sub LoadNew(FileName)
On Error GoTo Canceled
Open FileName For Input As #1
On Error GoTo PropperError

Input #1, totall

For n = 1 To totall
Add = FindBox
 Shape(Add).Used = True
  Input #1, Shape(Add).Name
  Input #1, Shape(Add).Colour
  For X = 1 To 6
  
   Input #1, foolish
   Shape(Add).Face(X).Reverse = foolish
   Input #1, foolish
   Shape(Add).Face(X).Used = foolish
   
   For Y = 1 To 4
    Input #1, Shape(Add).Face(X).Edge(Y)
   Next Y

   
  Next X

  For X = 1 To 8
   Input #1, Shape(Add).Corners(X).X
   Input #1, Shape(Add).Corners(X).Y
   Input #1, Shape(Add).Corners(X).z
  Next X


Next n
Close
Form1.Draw

Canceled:
Exit Sub
PropperError:
MsgBox "Agg, file failed to load!"
End Sub


Public Sub SaveAll(FileName, FileCheck)

On Error GoTo CouldntSave
If FileIsThere(FileName) = True And FileCheck = 1 Then
 X = MsgBox("File already exits. do you want to over-write it?", vbYesNo, "Wait...")
 If X = 7 Then Exit Sub
End If

Open FileName For Output As #1

totall = ShapeCount
Print #1, totall
For n = 1 To 100
 If Shape(n).Used = True Then
  Print #1, Shape(n).Name
  Print #1, Shape(n).Colour
  For X = 1 To 6
   Print #1, Shape(n).Face(X).Reverse
   Print #1, Shape(n).Face(X).Used
   For Y = 1 To 4
    Print #1, Shape(n).Face(X).Edge(Y)
   Next Y
  Next X
 
  For X = 1 To 8
   Print #1, Shape(n).Corners(X).X
   Print #1, Shape(n).Corners(X).Y
   Print #1, Shape(n).Corners(X).z
  Next X

End If

Next n
Close

Exit Sub
CouldntSave:
MsgBox "Error trying to save!"
Form1.mnuFileOptions(4).Enabled = False
End Sub


Public Function ShapeCount()
ShapeCount = 0
For n = 1 To 100
 If Shape(n).Used = True Then
  ShapeCount = ShapeCount + 1
 End If
 Next n
End Function

Public Function FileIsThere(FileName) As Boolean
FileIsThere = False
On Error GoTo NotThere
Open FileName For Input As #1
Close
FileIsThere = True
NotThere:
End Function

