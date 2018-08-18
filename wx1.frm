VERSION 5.00
Begin VB.Form wx1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "一元一次方程"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   2775
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "wx1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleMode       =   0  'User
   ScaleWidth      =   3544.061
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton mBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1440
      Width           =   1096
   End
   Begin VB.CommandButton sx1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "求解"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1440
      Width           =   1096
   End
   Begin VB.TextBox tx1b 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1680
      TabIndex        =   1
      Text            =   "b"
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox tx1a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   240
      TabIndex        =   0
      Text            =   "a"
      Top             =   240
      Width           =   855
   End
   Begin VB.Label rx1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x = "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label lx1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "x = "
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1200
      TabIndex        =   4
      Top             =   270
      Width           =   375
   End
   Begin VB.Menu mAbout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "wx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ml = "wx1"
End Sub

Private Sub mAbout_Click()
wAbout.Show 1
End Sub


Private Sub mBack_Click()
Unload Me
End Sub


Private Sub sx1_Click()
rx1.Caption = "x = "
If tx1a.Text = "a" Or tx1a.Text = "" Then
rx1.ForeColor = &HFF&
rx1.Caption = "请输入 a 的值"
ElseIf tx1b.Text = "b" Or tx1b.Text = "" Then
rx1.ForeColor = &HFF&
rx1.Caption = "请输入 b 的值"
ElseIf tx1a.Text = "0" Then
rx1.ForeColor = &HFF&
rx1.Caption = "系数 a 不能为零"
Else
rx1.ForeColor = &H80000012
rx1.Caption = rx1.Caption + Str(Val(tx1b.Text) / Val(tx1a.Text))
End If
End Sub
Private Sub txa_Change(Index As Integer)

End Sub


Private Sub tx1a_Click()
tx1a.Text = ""
End Sub


Private Sub tx1b_Click()
tx1b.Text = ""
End Sub


