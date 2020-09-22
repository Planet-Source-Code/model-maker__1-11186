Attribute VB_Name = "Module1"
Public Type coord
 X As Integer
 Y As Integer
 z As Integer
 MouseOver As Boolean
End Type


Public Type Faces
 Edge(4) As Byte
 Used As Boolean
 Reverse As Boolean
End Type

Public Type Shape
 Corners(8) As coord
 Face(6) As Faces
 Used As Boolean
 Name As String
 Colour As Double
End Type


Public Nifty(100) As Byte
Public Xmouse As Integer, Ymouse As Integer
Public Selected As Byte

Public Shape(100) As Shape
Public CopyBoard As Shape
Public Rotated As Shape
Public UndoBoard As Shape

Public Zoom As Single
Public Yofs, Xofs

Public Sine(0 To 361) As Double
Public Cosine(0 To 361) As Double

Public Const DisttoSelect = 4


Public Const PI = 3.14159265358979


Public Function CountBox()
CountBox = 0
For n = 1 To 100
 If Shape(n).Used = True Then
  CountBox = CountBox + 1
 End If
Next n
End Function

Public Function CountFace()
CountFace = 0
For n = 1 To 100
 If Shape(n).Used = True Then
  For m = 1 To 6
   If Shape(n).Face(m).Used = True Then
    CountFace = CountFace + 1
   End If
  Next m
 End If
Next n
End Function


Public Function FindBox()
For n = 1 To 100
 If Shape(n).Used = False Then
  FindBox = n: Exit Function
 End If
Next n
End Function

Public Sub DrawFromTop()
Dim x1 As Integer, z1 As Integer, x2  As Integer, z2 As Integer

'Form1.Caption = Selected
Form1.Names.Clear
For n = 1 To 100
 If Shape(n).Used = True Then
 Form1.Names.AddItem Shape(n).Name
   Nifty(Form1.Names.ListCount) = n '###############
 
For Fc = 1 To 6
  CenX = 0
  CenY = 0
  For m = 1 To 4
   If m = 4 Then
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).X
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).z
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).X
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).z
    Else
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).X
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).z
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).X
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).z
    End If
        

      
      x1 = x1 * Zoom + Xofs
      x2 = x2 * Zoom + Xofs
      z1 = z1 * Zoom + Yofs
      z2 = z2 * Zoom + Yofs
      
    CenX = CenX + x1
    CenY = CenY + z1
      
    If n = Selected Then
     If Form1.func(3) = False Then Form1.View.Circle (x1, z1), 3
     Form1.View.Line (x1 + 1, z1)-(x2 + 1, z2)
     Form1.View.Line (x1, z1 + 1)-(x2, z2 + 1)
    End If
    
    
    Form1.View.Line (x1, z1)-(x2, z2), Shape(n).Colour
  Next m

If n = Selected And Form1.func(3) = True Then Form1.View.Circle (CenX / 4, CenY / 4), 3

Next Fc


  
 End If
Next n

End Sub


Public Sub DrawFromSide()
Dim x1 As Integer, z1 As Integer, x2  As Integer, z2 As Integer

'Form1.Caption = Selected
Form1.Names.Clear
For n = 1 To 100
 If Shape(n).Used = True Then
 Form1.Names.AddItem Shape(n).Name
   Nifty(Form1.Names.ListCount) = n '###############
        
For Fc = 1 To 6
  CenX = 0
  CenY = 0
  For m = 1 To 4
   If m = 4 Then
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).X
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).Y
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).X
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).Y
    Else
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).X
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).Y
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).X
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).Y
    End If
        
      x1 = x1 * Zoom + Xofs
      x2 = x2 * Zoom + Xofs
      z1 = z1 * Zoom + Yofs
      z2 = z2 * Zoom + Yofs
        
    CenX = CenX + x1
    CenY = CenY + z1
    
    If n = Selected Then
     If Form1.func(3) = False Then Form1.View.Circle (x1, z1), 3
     Form1.View.Line (x1 + 1, z1)-(x2 + 1, z2)
     Form1.View.Line (x1, z1 + 1)-(x2, z2 + 1)
    End If
    
    Form1.View.Line (x1, z1)-(x2, z2), Shape(n).Colour
  Next m
  If n = Selected And Form1.func(3) = True Then Form1.View.Circle (CenX / 4, CenY / 4), 3

Next Fc
 End If
Next n

End Sub

Public Sub FillCube(Add)

 Shape(Add).Face(1).Edge(1) = 1
 Shape(Add).Face(1).Edge(2) = 2
 Shape(Add).Face(1).Edge(3) = 3
 Shape(Add).Face(1).Edge(4) = 4

 Shape(Add).Face(2).Edge(1) = 8
 Shape(Add).Face(2).Edge(2) = 7
 Shape(Add).Face(2).Edge(3) = 6
 Shape(Add).Face(2).Edge(4) = 5

 Shape(Add).Face(3).Edge(1) = 5
 Shape(Add).Face(3).Edge(2) = 6
 Shape(Add).Face(3).Edge(3) = 2
 Shape(Add).Face(3).Edge(4) = 1

 Shape(Add).Face(4).Edge(1) = 7
 Shape(Add).Face(4).Edge(2) = 8
 Shape(Add).Face(4).Edge(3) = 4
 Shape(Add).Face(4).Edge(4) = 3

 Shape(Add).Face(5).Edge(1) = 1
 Shape(Add).Face(5).Edge(2) = 4
 Shape(Add).Face(5).Edge(3) = 8
 Shape(Add).Face(5).Edge(4) = 5
 
 Shape(Add).Face(6).Edge(1) = 6
 Shape(Add).Face(6).Edge(2) = 7
 Shape(Add).Face(6).Edge(3) = 3
 Shape(Add).Face(6).Edge(4) = 2

 For n = 1 To 6
  Shape(Add).Face(n).Used = True
 Next n

End Sub



Public Sub DrawFromFront()
Dim x1 As Integer, z1 As Integer, x2  As Integer, z2 As Integer

'Form1.Caption = Selected
Form1.Names.Clear
For n = 1 To 100
 If Shape(n).Used = True Then
 Form1.Names.AddItem Shape(n).Name
 Nifty(Form1.Names.ListCount) = n '##############
        
For Fc = 1 To 6
  CenX = 0
  CenY = 0
  For m = 1 To 4
   If m = 4 Then
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).z
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(4)).Y
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).z
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(1)).Y
    Else
     x1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).z
     z1 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m)).Y
     x2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).z
     z2 = Shape(n).Corners(Shape(n).Face(Fc).Edge(m + 1)).Y
    End If
        
      x1 = x1 * Zoom + Xofs
      x2 = x2 * Zoom + Xofs
      z1 = z1 * Zoom + Yofs
      z2 = z2 * Zoom + Yofs
      
    CenX = CenX + x1
    CenY = CenY + z1
        
    If n = Selected Then
     If Form1.func(3) = False Then Form1.View.Circle (x1, z1), 3
     Form1.View.Line (x1 + 1, z1)-(x2 + 1, z2)
     Form1.View.Line (x1, z1 + 1)-(x2, z2 + 1)
    End If
    
    Form1.View.Line (x1, z1)-(x2, z2), Shape(n).Colour
  Next m
  If n = Selected And Form1.func(3) = True Then Form1.View.Circle (CenX / 4, CenY / 4), 3

Next Fc
 End If
Next n

End Sub


Public Sub ClearShapes()
  For n = 1 To 100
   Shape(n).Used = False
  Next n
End Sub

