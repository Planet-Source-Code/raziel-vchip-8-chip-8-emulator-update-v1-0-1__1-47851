Attribute VB_Name = "EmuMod"
Option Explicit
'DirecX Decs
Public DDraw     As clsDirectDraw7
Public bbuf      As clsDDSurface7
Public DInput    As clsDirectInput8
Public KeyHandle As clsDIKeyboard8

'General Decs
Public Ram(4100) As Byte
Public Const unible As Byte = &HF0
Public Const lnible As Byte = &HF
Public Const ubyte As Long = 65280
Public Const lbyte As Long = &HFF
Public Const llnible As Byte = &HF
Public Const lunible As Byte = &HF0
Public Const ulnible As Long = &HF00
Public Const uunible As Long = 61440
Public bRuning As Boolean
Public ssp  As Long
Public dFull As Boolean

'Mixer Vars
Public tCpu As Double
Public tGpu As Double
Public tIpu As Double
Public tSpu As Double

'Cpu Vars
Public Regs(15) As Byte
Public PC As Long
Public SP As Long
Public rI  As Long
Public Smem(15) As Long
Public dTimer As Long
Public rFlags(15) As Byte
'Input Vars
Public kp As Long
Public kdb(255) As Byte

'Sound Vars
Public sTimer As Long
Public bep As Boolean

'Gx    vars
Public Vram(143, 79) As Byte 'Update for 1.0.1
Public cFont(15, 5) As Byte
Public col(1) As Long
Public ds As Long
Public xrez As Long
Public yrez As Long
Public RECT As RECT
Public mtable(7) As Byte

Sub InitCpu()
'Gx init/reset
Dim i As Long
SP = 0
PC = 512
Regs(0) = 0
Regs(1) = 0
Regs(2) = 0
Regs(3) = 0
Regs(4) = 0
Regs(5) = 0
Regs(6) = 0
Regs(7) = 0
Regs(8) = 0
Regs(9) = 0
Regs(10) = 0
Regs(11) = 0
Regs(12) = 0
Regs(13) = 0
Regs(14) = 0
Regs(15) = 0
rI = 0
For i = 0 To 4095
Ram(i) = 0
Next i
End Sub
Sub InitGx()
'Gx init
If dFull Then
DDraw.Startup FrmRender.hWnd, 1024, 768, 32, True
Else
DDraw.Startup FrmMain.Picture1.hWnd, 512, 256, 32, False
End If
bbuf.Create DDraw.DDObj, 1025, 513
sm = FrmMain.sm.Value
sm2 = FrmMain.sm.Value
Dim i1 As Long, i2 As Long
col(0) = 0
col(1) = RGB(255, 255, 255)
For i1 = 0 To 127
For i2 = 0 To 63
Vram(i1, i2) = col(0)
Next i2
Next i1
bbuf.Surface.SetFillColor RGB(255, 255, 255)
bbuf.Surface.SetForeColor RGB(255, 255, 255)
End Sub
'Init Part
Sub InitEmu()
Dim c0 As Long
'load font
c0 = 0: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 9 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 15 * 16
c0 = 1: cFont(c0, 1) = 16: cFont(c0, 2) = 16: cFont(c0, 3) = 16: cFont(c0, 4) = 16: cFont(c0, 5) = 16
c0 = 2: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 1 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 8 * 16: cFont(c0, 5) = 15 * 16
c0 = 3: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 1 * 16: cFont(c0, 3) = 7 * 16: cFont(c0, 4) = 1 * 16: cFont(c0, 5) = 15 * 16
c0 = 4: cFont(c0, 1) = 9 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 1 * 16: cFont(c0, 5) = 1 * 16
c0 = 5: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 8 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 1 * 16: cFont(c0, 5) = 15 * 16
c0 = 6: cFont(c0, 1) = 8 * 16: cFont(c0, 2) = 8 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 15 * 16
c0 = 7: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 1 * 16: cFont(c0, 3) = 1 * 16: cFont(c0, 4) = 1 * 16: cFont(c0, 5) = 1 * 16
c0 = 8: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 15 * 16
c0 = 9: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 16: cFont(c0, 5) = 16
c0 = 10: cFont(c0, 1) = 6 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 15 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 9 * 16
c0 = 11: cFont(c0, 1) = 14 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 14 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 14 * 16
c0 = 12: cFont(c0, 1) = 6 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 8 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 6 * 16
c0 = 13: cFont(c0, 1) = 14 * 16: cFont(c0, 2) = 9 * 16: cFont(c0, 3) = 9 * 16: cFont(c0, 4) = 9 * 16: cFont(c0, 5) = 14 * 16
c0 = 14: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 8 * 16: cFont(c0, 3) = 14 * 16: cFont(c0, 4) = 8 * 16: cFont(c0, 5) = 15 * 16
c0 = 15: cFont(c0, 1) = 15 * 16: cFont(c0, 2) = 8 * 16: cFont(c0, 3) = 14 * 16: cFont(c0, 4) = 8 * 16: cFont(c0, 5) = 8 * 16
'init form
FrmMain.Picture1.AutoRedraw = False
FrmMain.BorderStyle = 0
FrmMain.Picture1.BorderStyle = 1
FrmMain.Picture1.ClipControls = False
FrmMain.Picture1.Visible = True
FrmMain.Picture1.ScaleMode = 1
FrmMain.Picture1.BackColor = RGB(0, 0, 0)
FrmMain.Picture1.Width = 512
FrmMain.Picture1.Height = 256
Set DDraw = New clsDirectDraw7
Set bbuf = New clsDDSurface7
Set DInput = New clsDirectInput8
Set KeyHandle = New clsDIKeyboard8
'Start Dinput
DInput.Startup FrmMain.hWnd
KeyHandle.Startup DInput, FrmMain.hWnd
'Setup needed vars
ds = 16
mtable(0) = 1
mtable(1) = 2
mtable(2) = 4
mtable(3) = 8
mtable(4) = 16
mtable(5) = 32
mtable(6) = 64
mtable(7) = 128
xrez = 64: yrez = 32
bRuning = False
ssp = 40
initkey
dFull = kdb(255) = 98
End Sub
Sub loadRom(str As String)
Dim i As Long
'Init everything
InitCpu
InitGx
initkey
'load a rom
Open str For Binary As #1
For i = 512 To LOF(1) + 511
Get #1, , Ram(i)
Next i
Close #1
bRuning = True
RunEmu
End Sub
'Mixer Part
'v 1.0.1 - dxtimer
'Fixed SmothMotion2
Sub RunEmu()
Do
DoEvents
RunCpu (ssp)
If GXH.sm2 Then If GXH.todraw = True Then GXH.smooth = True 'RunCpu (8)
ReDraw
chekIT
If dTimer Then dTimer = dTimer - 1: startbeep Else stopbeep
If sTimer Then sTimer = dTimer - 1
sleep 16
Loop While bRuning
End Sub
Sub sleep(milisecs As Double)
Dim mo As Double
mo = DDraw.DXObj.TickCount
milisecs = milisecs + mo
Do
DoEvents
Loop While DDraw.DXObj.TickCount < milisecs
End Sub
