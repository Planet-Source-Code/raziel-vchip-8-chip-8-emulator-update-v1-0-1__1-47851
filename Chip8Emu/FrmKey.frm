VERSION 5.00
Begin VB.Form FrmKey 
   Caption         =   "Configure input"
   ClientHeight    =   2430
   ClientLeft      =   2895
   ClientTop       =   2985
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   2430
   ScaleWidth      =   6840
   Begin VB.CheckBox fs 
      Caption         =   "FullSreen (1024x768 - no streach)"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   1560
      Width           =   6495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Go Back"
      Height          =   375
      Left            =   5280
      TabIndex        =   21
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK..Save It"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   1920
      Width           =   1455
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   15
      Left            =   5400
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   14
      Left            =   5400
      TabIndex        =   18
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   13
      Left            =   5400
      TabIndex        =   17
      Text            =   "Combo1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   12
      Left            =   5400
      TabIndex        =   16
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   11
      Left            =   3720
      TabIndex        =   14
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   9
      Left            =   3720
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   6
      Left            =   3720
      TabIndex        =   12
      Text            =   "Combo1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   3
      Left            =   3720
      TabIndex        =   11
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   0
      Left            =   2040
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   8
      Left            =   2040
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   5
      Left            =   2040
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   2
      Left            =   2040
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   10
      Left            =   360
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1200
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   4
      Left            =   360
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   480
      Width           =   1335
   End
   Begin VB.ComboBox Key 
      Height          =   315
      Index           =   1
      Left            =   360
      TabIndex        =   1
      Text            =   "Combo1"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "C    D    E   F"
      Height          =   1335
      Left            =   5160
      TabIndex        =   15
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "3    6    9   B"
      Height          =   1335
      Left            =   3480
      TabIndex        =   10
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label2 
      Caption         =   "2    5    8    0"
      Height          =   1335
      Left            =   1800
      TabIndex        =   5
      Top             =   120
      Width           =   135
   End
   Begin VB.Label Label1 
      Caption         =   "1    4    7   A"
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   135
   End
End
Attribute VB_Name = "FrmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim i As Long
For i = 0 To 254
kdb(i) = 99
Next i
For i = 0 To 15
kdb(Key(i).ListIndex) = i
Next i
Open App.Path & "/key.kdb" For Binary As #1
Put #1, , kdb
Close #1
Me.Hide
End Sub

Private Sub Command2_Click()
Me.Hide
Unload Me
End Sub

Private Sub Form_Load()
initkey
Dim i As Long, i2 As Long
Key(1).Clear
For i = 0 To 255
Key(1).AddItem KeyHandle.KeyName(i), i
Next i
Key(1).ListIndex = 0
Key(0).Clear
For i = 0 To 255
Key(0).List(i) = Key(1).List(i)
Next i
Key(0).ListIndex = 0
For i = 2 To 15
Key(i).Clear
For i2 = 0 To 255
Key(i).List(i2) = Key(i - 1).List(i2)
Next i2
Key(i).ListIndex = 0
Next i
Me.Show
For i = 0 To 255
If kdb(i) < 16 Then
Key(kdb(i)).ListIndex = i
End If
Next i
fs.Value = 99 - kdb(255)
End Sub

Private Sub fs_Click()
dFull = fs.Value
kdb(255) = 99 - Abs(dFull)
End Sub
