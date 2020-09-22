Attribute VB_Name = "Spin"
Sub Rotate(Angle1, Angle2, Angle3, NowOn, mx, my, mz)

For Cu = 1 To 8
   X = Shape(NowOn).Corners(Cu).X - mx
   Y = Shape(NowOn).Corners(Cu).Y - my
   Z = Shape(NowOn).Corners(Cu).Z - mz
   
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
   Rotated.Corners(Cu).X = Xrotated + mx
   Rotated.Corners(Cu).Y = Yrotated + my
   Rotated.Corners(Cu).Z = Zrotated + mz
   
 Next Cu
End Sub


Sub Make_LookUp()
 For i = 0 To 361
  Sine(i) = Sin(i / 180 * PI)
  Cosine(i) = Cos(i / 180 * PI)
 Next
End Sub


Function FindAngle(TXX, Tyy, PAngle)
 rxx = Cos(PAngle) * TXX - Sin(PAngle) * Tyy
 ryy = Sin(PAngle) * TXX + Cos(PAngle) * Tyy
 xx = rxx
 yy = ryy
 Pye = (22 / 7) * 18.05
 If yy = 0 Then
  If xx > 0 Then
   Angle = -90 / Pye
  Else
   Angle = -270 / Pye
  End If
 Else
  Angle = Atn(xx / yy)
 End If
 Angle = -Angle * Pye
 If yy > 0 Then Angle = Angle + 180
 If 0 > yy And xx < 0 Then Angle = Angle + 360
 If xx < 0 Then Angle = -360 + Angle
 Angle = Int(Angle)
 FindAngle = -Angle + 180
End Function

