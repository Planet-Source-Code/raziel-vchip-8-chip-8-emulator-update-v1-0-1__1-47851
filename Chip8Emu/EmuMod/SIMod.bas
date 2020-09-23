Attribute VB_Name = "SIMod"
'Input Part
'v 1.0.0 - DI8
Sub initkey()
Open App.Path & "/key.kdb" For Binary As #1
Get #1, , kdb
Close #1
End Sub
Sub chekIT()
kp = 99
KeyHandle.Check_Keyboard
End Sub
Sub KeyDown(ByVal lKey As Long, ByVal KeyName As String)
kp = kdb(lKey)
End Sub

'Sound Part
'v 0.0.1 - none
'I'l try to add DS8 support later ...
Sub startbeep()
If bep = False Then
'start beep
bep = True
End If
End Sub
Sub stopbeep()
'stop beep
bep = False
End Sub
