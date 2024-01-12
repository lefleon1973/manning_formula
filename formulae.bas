Attribute VB_Name = "Module2"
'Solve uniform flow manning equation
'Q flow rate in cubic meters per second
'n manning coefficient
'J channel slope
'D diameter of pipe in meters
Function yunicirc(Q As Double, n As Double, J As Double, D As Double) As Double
DY = 10 ^ (-7)
  'Q = Cells(rr, 8).Value
 ' J = J / 1000
 'k = Cells(4, 4).Value
 If Q > Atn(1) * D ^ 2 * (D / 4) ^ (2 / 3) * Sqr(J) / n Then
    'MsgBox ("Ανεπαρκής διατομή" + Chr(13) + Chr(10) + "Επιλέξτε μεγαλύτερη διάμετρο"), vbOKOnly
    yunicirc = CVErr(xlErrNum)
    Exit Function
End If
'D = D / 100
y = D / 2

Do
  If y < D Then f = 4 * Atn(Sqr(y / (D - y))) Else f = 8 * Atn(1)
  E = (f - Sin(f)) * D ^ 2 / 8
  R = D / 4 * (1 - Sin(f) / f)
  q1 = E * R ^ (2 / 3) * Sqr(J) / n - Q
  f = 4 * Atn(Sqr((y - DY) / (D - (y - DY))))
  E = (f - Sin(f)) * D ^ 2 / 8
  R = D / 4 * (1 - Sin(f) / f)
  Q2 = E * R ^ (2 / 3) * Sqr(J) / n - Q
  pr = (Q2 - q1) / (-DY)
  y = y - q1 / pr
  If y < 0 Then y = DY * 2
  If y > D Then y = D - DY
  i = i + 1
Loop Until Abs(q1) <= DY Or i = 100

yunicirc = y

End Function

Function yunitrapez(Q As Double, n As Double, J As Double, B As Double, z1 As Double, z2 As Double) As Double
DY = 10 ^ (-7)
  'Q = Cells(rr, 8).Value
'  J = J / 1000
 'k = Cells(4, 4).Value
 i = 0
 y = 0.001
DY = 30
ycr = 0.001
i = 0
Y1 = ycr
    e1 = B * ycr + 0.5 * (z1 + z2) * ycr ^ 2
    B1 = B + ycr * (z1 + z2)
    v1 = Q / e1
    K1 = v1 - Sqr(9.81 * e1 / B1)
    ycr = Y1 + DY
    E2 = B * ycr + 0.5 * (z1 + z2) * ycr ^ 2
    B2 = B + ycr * (z1 + z2)
    v2 = Q / E2
    K2 = v2 - Sqr(9.81 * E2 / B2)
    Y2 = ycr
    If K1 * K2 < 0 Then
    Do
    DY = DY / 2
    Y3 = Y1 + DY
    E3 = B * Y3 + 0.5 * (z1 + z2) * Y3 ^ 2
    B3 = B + Y3 * (z1 + z2)
    v3 = Q / E3
    k3 = v3 - Sqr(9.81 * E3 / B3)
    If K1 * k3 < 0 Then Y2 = Y3: K2 = k3 Else Y1 = Y3: K1 = k3
    Loop Until DY < 0.00000001
End If
ycr = (Y1 + Y2) / 2


DY = 30
Y1 = 0.001
y = Y1
    e1 = B * y + 0.5 * (z1 + z2) * y ^ 2
    P1 = B + y * (Sqr(1 + z1 ^ 2) + Sqr(1 + z2 ^ 2))
    R1 = e1 / P1
    K1 = Q - e1 * R1 ^ (2 / 3) * J ^ (1 / 2) / n
    y = y + DY
    Y2 = y
    E2 = B * y + 0.5 * (z1 + z2) * y ^ 2
    P2 = B + y * (Sqr(1 + z1 ^ 2) + Sqr(1 + z2 ^ 2))
    R2 = E2 / P2
    K2 = Q - E2 * R2 ^ (2 / 3) * J ^ (1 / 2) / n
    If K1 * K2 < 0 Then
    Do
    DY = DY / 2
    Y3 = Y1 + DY
    E3 = B * Y3 + 0.5 * (z1 + z2) * Y3 ^ 2
    P3 = B + Y3 * (Sqr(1 + z1 ^ 2) + Sqr(1 + z2 ^ 2))
    R3 = E3 / P3
    k3 = Q - E3 * R3 ^ (2 / 3) * J ^ (1 / 2) / n
    If K1 * k3 < 0 Then Y2 = Y3: K2 = k3 Else Y1 = Y3: K1 = k3
    
    Loop Until DY < 0.000000001
End If
y = (Y1 + Y2) / 2
'Cells(rr, 14).Value = y
yunitrapez = y
End Function

