Attribute VB_Name = "CPUMod"
Option Explicit
Dim memptr As Long, i1 As Long, i2 As Long, i3 As Long, i4 As Long, tmp1 As Long, tmp2 As Long, pci As Long
'Cpu   Part
'v 1.0.1 - Interpreter
'Some Bugfixes(rnd,8XY4,Wait Key)
Sub RunCpu(rc As Long)
For pci = 0 To rc
i1 = Ram(PC) \ 16: i2 = (Ram(PC) And lnible): PC = PC + 1: i3 = Ram(PC) \ 16: i4 = Ram(PC) And lnible: PC = PC + 1

Select Case i1
Case 0
Select Case i3
    Case 12 '   Scroll down N lines (Schip)
        
        For tmp2 = yrez - 1 - i4 To 0 Step -1
        For tmp1 = 0 To xrez - 1
        Vram(tmp1, tmp2 + i4) = Vram(tmp1, tmp2)
        Next tmp1
        Next tmp2
        
        For tmp2 = 0 To i4 - 1
        For tmp1 = 0 To xrez - 1
        Vram(tmp1, tmp2) = 0
        Next tmp1
        Next tmp2
        
    Case 14

    If i4 = 0 Then     '0        Erase the screen
    For tmp1 = 0 To yrez - 1
    For tmp2 = 0 To xrez - 1
    Vram(tmp2, tmp1) = col(0)
    Next tmp2
    Next tmp1
    ElseIf i4 = 14 Then 'E        Return from a CHIP-8 sub-routine
    SP = SP - 1
    PC = Smem(SP)
    End If
    Case 15
    
    Select Case i4
        Case 11 'B        Scroll 4 pixels right (Schip)
        For tmp2 = 0 To yrez - 1
        For tmp1 = xrez - 1 - xrez \ 32 To 0 Step -1
        Vram(tmp1 + xrez \ 32, tmp2) = Vram(tmp1, tmp2)
        Next tmp1
        Next tmp2
        
        For tmp2 = 0 To yrez - 1
        For tmp1 = 0 To xrez \ 32 - 1
        Vram(tmp1, tmp2) = 0
        Next tmp1
        Next tmp2
        
        Case 12 'C        Scroll 4 pixels left (Schip)
        For tmp2 = 0 To yrez - 1
        For tmp1 = xrez \ 32 To xrez - 1
        Vram(tmp1 - xrez \ 32, tmp2) = Vram(tmp1, tmp2)
        Next tmp1
        Next tmp2
        
        For tmp2 = 0 To yrez - 1
        For tmp1 = xrez - xrez \ 32 To xrez - 1
        Vram(tmp1, tmp2) = 0
        Next tmp1
        Next tmp2
        
        Case 13 'D        Quit the emulator (Schip)

        bRuning = False
        Case 14 'E        Set CHIP-8 graphic mode (Schip)

        xrez = 64: yrez = 32: ds = 16
        Case 15 'F        Set SCHIP graphic mode (Schip)

        xrez = 128: yrez = 64: ds = 8
    End Select
End Select

Case 1 'Jump to NNN

PC = i2 * 256 + i3 * 16 + i4
Case 2 'Call CHIP-8 sub-routine at NNN (16 successive calls max)

Smem(SP) = PC
SP = SP + 1
PC = i2 * 256 + i3 * 16 + i4
Case 3 'Skip next instruction if VX = KK

If Regs(i2) = i3 * 16 + i4 Then PC = PC + 2
Case 4 'Skip next instruction if VX <> KK

If Regs(i2) <> i3 * 16 + i4 Then PC = PC + 2
Case 5 'Skip next instruction if VX = VY

If Regs(i2) = Regs(i3) Then PC = PC + 2
Case 6 'VX = KK

Regs(i2) = i3 * 16 + i4
Case 7 'VX = VX + KK

Regs(i2) = (CLng(Regs(i2)) + i3 * 16 + i4) And 255
Case 8 'multi..

Select Case i4
    Case 0 'VX = VY
    Regs(i2) = Regs(i3)
    Case 1 'VX = VX OR VY
    Regs(i2) = Regs(i2) Or Regs(i3)
    Case 2 'VX = VX AND VY
    Regs(i2) = Regs(i2) And Regs(i3)
    Case 3 'VX = VX XOR VY
    Regs(i2) = Regs(i2) Xor Regs(i3)
    Case 4 'VX = VX + VY, VF = carry
    tmp1 = CLng(Regs(i2)) + Regs(i3)
    If tmp1 > 255 Then Regs(15) = 1: tmp1 = tmp1 - 256 Else Regs(15) = 0
    Regs(i2) = tmp1
    Case 5 'VX = VX - VY, VF = not borrow
    tmp2 = CLng(Regs(i2)) - Regs(i3)
    If tmp2 < 0 Then Regs(15) = 0: tmp2 = tmp2 + 256 Else Regs(15) = 1
    Regs(i2) = tmp2
    Case 6 'VX = VX SHR 1 (VX=VX/2), VF = carry
    Regs(15) = Regs(i2) And 1
    Regs(i2) = Regs(i2) \ 2
    Case 7 'VX = VY - VX, VF = not borrow
    tmp2 = CLng(Regs(i3)) - Regs(i2)
    If tmp2 <= 0 Then Regs(15) = 1: tmp2 = tmp2 + 256 Else Regs(15) = 0
    Regs(i2) = tmp2
    Case 14 'VX = VX SHL 1 (VX=VX*2), VF = carry
    tmp2 = Regs(i2) * 2
    If Regs(i2) > 256 Then Regs(15) = 1 Else Regs(15) = 0
    Regs(i2) = tmp2 And 255
End Select
Case 9 'Skip next instruction if VX != VY

If Regs(i2) <> Regs(i3) Then PC = PC + 2
Case 10 'I = NNN

rI = i2 * 256 + i3 * 16 + i4
Case 11 'Jump to NNN + V0

PC = i2 * 256 + i3 * 16 + i4 + Regs(0)
Case 12 'VX = Random number AND KK

Regs(i2) = (Rnd * 255) And (i3 * 16 + i4)
Case 13 'Draws a sprite at (VX,VY) starting at M(I). VF = collision.If N=0, draws the 16 x 16 sprite, else an 8 x N sprite.
DrawSprite Regs(i2), Regs(i3), rI, i4
Case 14

If i3 = 9 Then 'Skip next instruction if key VX pressed
    If kp = Regs(i2) Then PC = PC + 2
ElseIf i3 = 10 Then 'Skip next instruction if key VX not pressed
    If kp <> Regs(i2) Then PC = PC + 2
End If
Case 15

Select Case i3
    Case 0
    
    If i4 = 7 Then 'VX = Delay timer
        Regs(i2) = dTimer
    ElseIf i4 = 10 Then 'Waits a keypress and stores it in VX
    If kp = 99 Then PC = PC - 2: chekIT Else Regs(i2) = kp
    End If
    Case 1
    
    If i4 = 5 Then 'Delay timer = VX
        dTimer = Regs(i2)
    ElseIf i4 = 8 Then 'Sound timer = VX
        sTimer = Regs(i2)
    ElseIf i4 = 14 Then 'i = i + vx
        rI = rI + Regs(i2)
    End If
    
    Case 2 'I points to the 4 x 5 font sprite of hex char in VX
    rI = 8000 + Regs(i2)
    Case 3 'Store BCD representation of VX in M(I)...M(I+2)
    tmp2 = 100
    For tmp1 = rI To rI + 2
    Ram(tmp1) = (Regs(i2) \ tmp2) Mod 10
    tmp2 = tmp2 \ 10
    Next tmp1
    Case 5 'Save V0...VX in memory starting at M(I)
    For tmp1 = rI To rI + i2
    Ram(tmp1) = Regs(tmp1 - rI)
    Next tmp1
    Case 6 'Load V0...VX from memory starting at M(I)
    For tmp1 = rI To rI + i2
    Regs(tmp1 - rI) = Ram(tmp1)
    Next tmp1
    Case 7 'Save V0...VX (X<8) in the HP48 flags (Schip)
    For tmp1 = 0 To i2
    rFlags(tmp1) = Ram(tmp1)
    Next tmp1
    Case 8 'Load V0...VX (X<8) from the HP48 flags (Schip)
    For tmp1 = 0 To i2
    Ram(tmp1) = rFlags(tmp1)
    Next tmp1
End Select
End Select
Next pci
End Sub
Sub SaveStage(Id As Long)
Open App.Path & "\" & FrmMain.CommonDialog1.FileTitle & ".ss." & Id & ".sst" For Binary As #1
Put #1, , Ram
Put #1, , Regs
Put #1, , PC
Put #1, , SP
Put #1, , Smem
Put #1, , sTimer
Put #1, , dTimer
Put #1, , rFlags
Put #1, , Vram
Put #1, , xrez
Put #1, , xrez
Close #1
End Sub
Sub LoadStage(Id As Long)
Open App.Path & "\" & FrmMain.CommonDialog1.FileTitle & ".ss." & Id & ".sst" For Binary As #1
Get #1, , Ram
Get #1, , Regs
Get #1, , PC
Get #1, , SP
Get #1, , Smem
Get #1, , sTimer
Get #1, , dTimer
Get #1, , rFlags
Get #1, , Vram
Get #1, , xrez
Get #1, , xrez
Close #1
End Sub
