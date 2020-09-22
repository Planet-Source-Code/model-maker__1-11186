Attribute VB_Name = "TextureMapper"
Sub Poly()




Dim cLeft As Byte, cRight As Byte, cTop As Byte, cBottom As Byte

For n = 1 To 4
 If ner(n).X < 0 Then cLeft = cLeft + 1
 If ner(n).X > Form1.View.ScaleWidth Then cRight = cRight + 1
 If ner(n).Y < 0 Then cTop = cTop + 1
 If ner(n).Y > Form1.View.ScaleHeight Then cBottom = cBottom + 1
Next n

If cLeft = 4 Or cRight = 4 Or cTop = 4 Or cBottom = 4 Then
 Exit Sub
End If




Dim startstop(1000, 7)



tx = 50    'Size of texture
ty = 50    'Size of texture

For n = 1 To 4
 If n = 4 Then
  X1 = ner(n).X: Y1 = ner(n).Y
  X2 = ner(1).X: Y2 = ner(1).Y
 Else
  X1 = ner(n).X: Y1 = ner(n).Y
  X2 = ner(n + 1).X: Y2 = ner(n + 1).Y
 End If
 GoSub ScanTheLine
Next n

Exit Sub

ScanTheLine:
 Tag = X1: If Y2 - Y1 = 0 Then Return
 Shif = (X2 - X1) / (Y2 - Y1)
 If n < 3 Then LineTop = Y1
 If n > 2 Then LineTop = Y2
 pn = 1: If Y1 > Y2 Then pn = -1
 lineheight = Abs(Y1 - Y2)
 SmallValue = (tx / lineheight)
 Shif = Shif * pn
 For m = Y1 To Y2 Step pn
 If m > 0 And m < 1000 Then
 posfoo = SmallValue * Abs(m - LineTop)
   If startstop(m, 1) = 0 Then
     startstop(m, 1) = 1
     startstop(m, 2) = Int(Tag)
     startstop(m, 3) = n
     startstop(m, 4) = posfoo
    Else
    ' xx1 = startstop(m, 2)
    ' xx2 = Tag
    'View.Line (xx1, m)-(xx2, m), RGB(Rnd * 255, Rnd * 255, Rnd * 255)
 
    GoSub DrawATexturedLine
   End If
   Tag = Tag + Shif
  End If
 Next m
Return



DrawATexturedLine:
 Dim lr As Integer
 xx1 = startstop(m, 2): xx2 = Tag
 lr = 1: If xx1 > xx2 Then lr = -1
 
 linelength = Abs(xx1 - xx2)
 If linelength = 0 Then Return
 
 Select Case startstop(m, 3)
  Case 1: v1 = 0: u1 = startstop(m, 4)
  Case 2: v1 = startstop(m, 4): u1 = tx
  Case 3: v1 = ty: u1 = startstop(m, 4)
  Case 4: v1 = startstop(m, 4): u1 = 0
 End Select
 Select Case n
  Case 1: v2 = 0: u2 = posfoo
  Case 2: v2 = posfoo: u2 = tx
  Case 3: v2 = ty: u2 = posfoo
  Case 4: v2 = posfoo: u2 = 0
 End Select
 u = u1: v = v1
 mlr = (u2 - u1) / linelength
 mud = (v2 - v1) / linelength
 

 For s = xx1 To xx2 Step 1
  u = u + mlr
  v = v + mud
  If u > tx Then u = tx
  If u < 0 Then u = 0
  If v > ty Then v = ty
  If v < 0 Then v = 0
  'red = tex(u - 0.5, v - 0.5).red
  'green = tex(u - 0.5, v - 0.5).green
  'blue = tex(u - 0.5, v - 0.5).blue
  col = tex(u - 0.5, v - 0.5)
'  MsgBox red & " " & green & " " & blue
  Form1.View.PSet (s, m), col
  
 Next s
 
 For s = xx1 To xx2 Step -1
  u = u + mlr
  v = v + mud
  If u > tx Then u = tx
  If u < 0 Then u = 0
  If v > ty Then v = ty
  If v < 0 Then v = 0
 ' red = tex(u - 0.5, v - 0.5).red
 ' green = tex(u - 0.5, v - 0.5).green
 ' blue = tex(u - 0.5, v - 0.5).blue
  col = tex(u - 0.5, v - 0.5)
  Form1.View.PSet (s, m), col
 Next s
Return
End Sub

