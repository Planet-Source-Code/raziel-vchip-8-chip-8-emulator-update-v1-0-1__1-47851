Attribute VB_Name = "GXH"
Option Explicit
Public sm As Boolean
Public sm2 As Boolean
Public smooth As Boolean, todraw As Boolean
'GXH - v 1.0.1
'Fixed Y Calc
Sub DrawSprite(ByVal X As Long, ByVal Y As Long, ByVal memptr As Long, ByVal n As Long)
'Draw  sprite to vram
Dim xend As Long, yend As Long, i1 As Long, i2 As Long, mi As Long, xr2 As Long, yr2 As Long
yr2 = yrez - 1: xr2 = xrez - 1
X = X Mod xrez: Y = Y Mod yrez
mi = 7
Regs(15) = 0
If n = 0 Then xend = X + 15: yend = Y + 15 Else xend = X + 7: yend = Y + n - 1
If memptr > 7999 Then
Ram(4095) = cFont(memptr - 8000, 0)
Ram(4096) = cFont(memptr - 8000, 1)
Ram(4097) = cFont(memptr - 8000, 2)
Ram(4098) = cFont(memptr - 8000, 3)
Ram(4099) = cFont(memptr - 8000, 4)
Ram(4100) = cFont(memptr - 8000, 5)
memptr = 4096
End If
For i1 = Y To yend
For i2 = X To xend
If ((Ram(memptr) And mtable(mi)) \ mtable(mi) Xor Vram(i2, i1)) = 0 And Vram(i2, i1) = 1 Then Regs(15) = 1
Vram(i2, i1) = (Ram(memptr) And mtable(mi)) \ mtable(mi) Xor Vram(i2, i1)
memptr = memptr + Abs(mi = 0)
mi = (mi - 1) And 7
Next i2
Next i1
If sm Then smooth = Regs(15) = 0: If smooth And todraw Then ReDraw: todraw = False
End Sub
Sub ReDraw()
'Convert and draw Vram to Display
If sm Then If smooth = False Then todraw = True: Exit Sub Else todraw = False:
Dim i1 As Long, i2 As Long
bbuf.BltColorFill 0
For i1 = 0 To yrez - 1
For i2 = 0 To xrez - 1
If Vram(i2, i1) Then
bbuf.Surface.DrawBox (i2 * ds), (i1 * ds), (i2 * ds) + ds, (i1 * ds) + ds
End If
Next i2
Next i1
If dFull Then
RECT = bbuf.SurfaceRect
RECT.Top = RECT.Top + 128
RECT.Bottom = RECT.Bottom + 128
DDraw.Screen.Blt RECT, bbuf.Surface, bbuf.SurfaceRect
Else
DDraw.Draw bbuf, bbuf.SurfaceRect
End If
End Sub
