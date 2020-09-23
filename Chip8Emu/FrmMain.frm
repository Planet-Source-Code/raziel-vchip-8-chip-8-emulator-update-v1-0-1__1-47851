VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmMain 
   AutoRedraw      =   -1  'True
   Caption         =   "VChip8"
   ClientHeight    =   4380
   ClientLeft      =   1605
   ClientTop       =   1770
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   ScaleHeight     =   292
   ScaleMode       =   0  'User
   ScaleWidth      =   525
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8040
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   3600
      Max             =   400
      Min             =   1
      TabIndex        =   3
      Top             =   4080
      Value           =   40
      Width           =   4215
   End
   Begin VB.CheckBox sm2 
      Caption         =   "SmothMotion V2"
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.CheckBox sm 
      Caption         =   "SmothMotion"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Value           =   1  'Checked
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   3825
      Left            =   120
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   511
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7665
   End
   Begin VB.Menu fi 
      Caption         =   "File"
      Begin VB.Menu op 
         Caption         =   "Open"
         Shortcut        =   {F1}
      End
      Begin VB.Menu opt 
         Caption         =   "Options"
         Shortcut        =   {F2}
      End
      Begin VB.Menu ss 
         Caption         =   "SaveStage"
         Shortcut        =   {F3}
      End
      Begin VB.Menu ls 
         Caption         =   "LoadStage"
         Shortcut        =   {F4}
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu Ab 
      Caption         =   "About"
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Ab_Click()
MsgBox "                 Vchip 8           " & vbNewLine & _
       "A chip 8 emulator writen in Vb" & vbNewLine & _
       "Coder : Raziel" & vbNewLine & _
       "Thanks to David WINTER (HPMANIAC)", vbOKOnly, "About Vchip 8"
End Sub

Private Sub ext_Click()
End
End Sub

Private Sub Form_DblClick()
End
End Sub

Private Sub Form_Load()
InitEmu
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

Private Sub HScroll1_Change()
ssp = HScroll1.Value
End Sub

Private Sub ls_Click()
LoadStage InputBox("Give the id of the SaveStage to load", "")
End Sub

Private Sub op_Click()
'On Error GoTo 12
CommonDialog1.CancelError = False
CommonDialog1.DialogTitle = "Open Chip8 Rom file"
CommonDialog1.Filter = "Chip8 Roms(*.*)|*.*"
CommonDialog1.Filename = ""
CommonDialog1.ShowOpen
If Len(CommonDialog1.Filename) > 1 Then loadRom CommonDialog1.Filename
'12:
End Sub



Private Sub opt_Click()
FrmKey.Show
End Sub

Private Sub sm_Click()
GXH.sm = sm.Value
sm2.Enabled = GXH.sm
End Sub

Private Sub sm2_Click()
GXH.sm2 = sm2.Value
End Sub

Private Sub ss_Click()
SaveStage InputBox("Give an id for this SaveStage", "")
End Sub
