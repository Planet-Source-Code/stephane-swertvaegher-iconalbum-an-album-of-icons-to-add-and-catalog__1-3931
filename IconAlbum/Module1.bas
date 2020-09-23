Attribute VB_Name = "Module1"
Public xx%, Idx%, IconIdx%
Public Temp$, IBpath$, inp$, Mess$

Public Sub ColForm(Obj As Object, R%, G%, B%, Step%)
Dim R1%, G1%, B1%, R2%, G2%, B2%
'Best effect with borderstyle = 0
'Obj   = Form or Picture-box
'R%    = color for red-component
'G%    = color for green-component
'B%    = color for blue-component
'Step% = the step to darken and lighten the border
Obj.ScaleMode = 3 'pixels
Obj.AutoRedraw = True 'very important !
Obj.BackColor = RGB(R%, G%, B%)
R1% = R% + Step%: If R1% > 255 Then R1% = 255
G1% = G% + Step%: If G1% > 255 Then G1% = 255
B1% = B% + Step%: If B1% > 255 Then B1% = 255
R2% = R% - Step%: If R2% < 0 Then R2% = 0
G2% = G% - Step%: If G2% < 0 Then G2% = 0
B2% = B% - Step%: If B2% < 0 Then B2% = 0
Obj.Line (2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R1%, G1%, B1%), B
Obj.Line (Obj.ScaleWidth - 2, 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 1), RGB(R2%, G2%, B2%)
Obj.Line (1, Obj.ScaleHeight - 2)-(Obj.ScaleWidth - 2, Obj.ScaleHeight - 2), RGB(R2%, G2%, B2%)
Obj.Line (5, 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R2%, G2%, B2%), B
Obj.Line (Obj.ScaleWidth - 5, 6)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 4), RGB(R1%, G1%, B1%)
Obj.Line (5, Obj.ScaleHeight - 5)-(Obj.ScaleWidth - 5, Obj.ScaleHeight - 5), RGB(R1%, G1%, B1%)
End Sub

