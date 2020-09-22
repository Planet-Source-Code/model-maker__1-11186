Attribute VB_Name = "DrawAModel"
Sub DrawModel()
Dim nn As Byte
Dim Xofs As Integer, Yofs As Integer, Zeye As Integer
Dim MTD As Byte


MTD = Entity.Model

Xofs = Form1.View.ScaleWidth / 2
Yofs = Form1.View.ScaleHeight / 2
Zeye = 800

Dim FaceOrder(250, 2) As Single

Call Rotate(Entity.Angle.X, Entity.Angle.Y, Entity.Angle.Z)


For n = 1 To Model(MTD).TotalEdges
 FaceOrder(n, 1) = n
 FaceOrder(n, 2) = Rotated(Model(MTD).Edge(n, 1)).Z
 For m = 2 To 4
  FaceOrder(n, 2) = FaceOrder(n, 2) + Rotated(Model(MTD).Edge(n, m)).Z
 Next m
Next n

Dim startdis As Integer
For n = 1 To Model(MTD).TotalEdges
 startdis = FaceOrder(n, 2)
 lok = n
 For m = n To Model(MTD).TotalEdges
  If FaceOrder(m, 2) < startdis Then
   lok = m
   startdis = FaceOrder(m, 2)
  End If
 Next m
 temp = FaceOrder(n, 2)
 FaceOrder(n, 2) = FaceOrder(lok, 2)
 FaceOrder(lok, 2) = temp
 temp = FaceOrder(n, 1)
 FaceOrder(n, 1) = FaceOrder(lok, 1)
 FaceOrder(lok, 1) = temp
Next n




For EachFace = 1 To Model(MTD).TotalEdges

 nn = FaceOrder(EachFace, 1)

 For m = 1 To 4
  X = Rotated(Model(MTD).Edge(nn, m)).X + Entity.Origin.X
  Y = Rotated(Model(MTD).Edge(nn, m)).Y + Entity.Origin.Y
  Z = Rotated(Model(MTD).Edge(nn, m)).Z + Entity.Origin.Z
  ner(m).X = Xofs + Int(X * (Zeye / (Zeye - Z)))
  ner(m).Y = Yofs + Int(Y * (Zeye / (Zeye - Z)))
  
  
 Next m

 xx1 = ner(1).X
 xx2 = ner(2).X
 xx3 = ner(3).X
 yy1 = ner(1).Y
 yy2 = ner(2).Y
 yy3 = ner(3).Y
 Normal = ((yy1 - yy3) * (xx2 - xx1) - (xx1 - xx3) * (yy2 - yy1))

  If Normal < 0 Then

  For m = 1 To 4
   If m = 4 Then
    X1 = ner(1).X
    Y1 = ner(1).Y
    X2 = ner(m).X
    Y2 = ner(m).Y
   Else
    X1 = ner(m).X
    Y1 = ner(m).Y
    X2 = ner(m + 1).X
    Y2 = ner(m + 1).Y
   End If
    Form1.View.Line (X1, Y1)-(X2, Y2)
    
  Next m
   If Form1.opt2 = 1 Then Poly

 End If
 
Next EachFace
End Sub

