Attribute VB_Name = "Suff_Nonsense"
Public Type Colour
 red As Byte
 green As Byte
 blue As Byte
End Type

Public Type coord
 X As Integer
 Y As Integer
 Z As Integer
End Type

Public Type StrtStop
 Finish As Integer
 Used As Boolean
 Edge As Byte
 TexLock As Single
End Type

Public Rotated(200) As coord

Type Model
 TotalPoints As Integer
 TotalEdges As Integer
 Points(200) As coord
 EdgeUsed(200) As Boolean
 Edge(200, 4) As Integer
End Type

Public Type Objectzz
 Origin As coord
 Angle As coord
 Model As Byte
End Type


Public Type Worldz
 Origin As coord
 Used As Boolean
End Type


Public Model(10) As Model
Public World(100) As Worldz
Public ner(4) As coord


Public Entity As Objectzz



Public Sine(0 To 361) As Double
Public Cosine(0 To 361) As Double
Public tex(150, 150) As Long
Public Const PI = 3.14159265358979


Public Sub Setup()

ModelNumber = 1
Open Form1.FileName.FileName For Input As #1
Input #1, Model(ModelNumber).TotalPoints
Input #1, Model(ModelNumber).TotalEdges

 Model(ModelNumber).TotalPoints = Model(ModelNumber).TotalPoints + 1
 Model(ModelNumber).TotalEdges = Model(ModelNumber).TotalEdges / 5


For n = 1 To Model(ModelNumber).TotalPoints
 Input #1, Model(ModelNumber).Points(n).X
 Input #1, Model(ModelNumber).Points(n).Y
 Input #1, Model(ModelNumber).Points(n).Z
Next n

For n = 1 To Model(ModelNumber).TotalEdges
 Input #1, edgeypopo
 For m = 1 To 4
  Input #1, Model(ModelNumber).Edge(n, m)
  Model(ModelNumber).Edge(n, m) = Model(ModelNumber).Edge(n, m) + 1
 Next m
Next n


Close

End Sub

Sub Random_World()
For n = 1 To 10
 World(n).Used = True
 World(n).Origin.X = Rnd * 1000
 World(n).Origin.Z = Rnd * 1000
Next n
End Sub

Sub Make_LookUp()
 For i = 0 To 361
  Sine(i) = Sin(i / 180 * PI)
  Cosine(i) = Cos(i / 180 * PI)
 Next
End Sub

Sub Rotate(Angle1, Angle2, Angle3)
Dim MTD As Byte
MTD = Entity.Model

For Cu = 1 To Model(MTD).TotalPoints
   X = Model(MTD).Points(Cu).X
   Y = Model(MTD).Points(Cu).Y
   Z = Model(MTD).Points(Cu).Z
   
   Xrotated = X
   Yrotated = Cosine(Angle1) * Y - Sine(Angle1) * Z
   Zrotated = Sine(Angle1) * Y + Cosine(Angle1) * Z
   X = Xrotated
   Y = Yrotated
   Z = Zrotated

   Xrotated = Cosine(Angle2) * X - Sine(Angle2) * Z
   Yrotated = Y
   Zrotated = Sine(Angle2) * X + Cosine(Angle2) * Z
   X = Xrotated
   Y = Yrotated
   Z = Zrotated

   Xrotated = Cosine(Angle3) * X - Sine(Angle3) * Y
   Yrotated = Sine(Angle3) * X + Cosine(Angle3) * Y
   Zrotated = Z
   Rotated(Cu).X = Xrotated
   Rotated(Cu).Y = Yrotated
   Rotated(Cu).Z = Zrotated
   
 Next Cu
End Sub

Sub GetTexture()



'Open "c:\texta.txt" For Input As #1
For n = 0 To 120
 For m = 0 To 120
  'Input #1, X
   tex(n, m) = RGB(100, 100 + Rnd * 155, 100)
   'QBColor(X)
 Next m
Next n

 For m = 0 To 120
   tex(24, m) = RGB(255, 255, 255)
   tex(25, m) = RGB(255, 255, 255)
   tex(26, m) = RGB(255, 255, 255)
 Next
 For m = 0 To 120
   tex(m, 24) = RGB(255, 255, 255)
   tex(m, 25) = RGB(255, 255, 255)
   tex(m, 26) = RGB(255, 255, 255)
 Next m


'Close
End Sub




