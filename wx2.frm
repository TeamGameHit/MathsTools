VERSION 5.00
Begin VB.Form wx2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ò»Ôª¶þ´Î·½³Ì"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   5325
   FillColor       =   &H00C0C0C0&
   LinkTopic       =   "wx2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleMode       =   0  'User
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton mBack 
      BackColor       =   &H00FFFFFF&
      Caption         =   "·µ»Ø"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1440
      Width           =   1096
   End
   Begin VB.CommandButton sx2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Çó½â"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      TabIndex        =   3
      Top             =   1440
      Width           =   1096
   End
   Begin VB.TextBox tx2c 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3480
      TabIndex        =   2
      Text            =   "c"
      Top             =   240
      Width           =   1092
   End
   Begin VB.TextBox tx2b 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1920
      TabIndex        =   1
      Text            =   "b"
      Top             =   240
      Width           =   1092
   End
   Begin VB.TextBox tx2a 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
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
      Width           =   1092
   End
   Begin VB.Label delta 
      BackColor       =   &H00FFFFFF&
      Caption         =   "¦¤="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   11
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Label lx24 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2790
      TabIndex        =   10
      Top             =   1000
      Width           =   135
   End
   Begin VB.Label lx23 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1000
      Width           =   135
   End
   Begin VB.Label lx26 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x = "
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   8
      Top             =   840
      Width           =   500
   End
   Begin VB.Label lx25 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x = "
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   500
   End
   Begin VB.Label lx22 
      BackColor       =   &H00FFFFFF&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   135
   End
   Begin VB.Label lx21 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x  +                     x +                     =0 "
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   270
      Width           =   3615
   End
   Begin VB.Label rx22 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   13
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label rx21 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   12
      Top             =   840
      Width           =   1695
   End
   Begin VB.Menu mAbout 
      Caption         =   "¹ØÓÚ"
   End
End
Attribute VB_Name = "wx2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lx1_Click()

End Sub

Private Sub lx2_Click()

End Sub


Private Sub rx2_Click(Index As Integer)

End Sub

Private Sub Command1_Click()
Print tx2a.Text, tx2b.Text, tx2c.Text, delta.Caption
Print ((-1) * Val(tx2b.Text) + Sqr(delta.Caption)) / 2 * Val(tx2a.Text)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub mAbout_Click()
wAbout.Show 1
End Sub

Private Sub mBack_Click()
Unload Me
End Sub

Private Sub sx2_Click()
delta.Caption = Str((Val(tx2b.Text)) ^ 2 - 4 * Val(tx2a.Text) * Val(tx2c.Text))
If Val(delta.Caption) > 0 Then
rx21.Caption = Str(((-1) * Val(tx2b.Text) + Sqr(Val(delta.Caption))) / 2 * Val(tx2a.Text))
rx22.Caption = Str(((-1) * Val(tx2b.Text) - Sqr(Val(delta.Caption))) / 2 * Val(tx2a.Text))
lx23.Visible = True
lx24.Visible = True
    lx24.Left = 2790
lx25.Visible = True
lx26.Visible = True
    lx26.Left = 2640
rx21.Left = 720
ElseIf Val(delta.Caption) = 0 Then
rx21.Caption = Str(((-1) * Val(tx2b.Text) + Sqr(Val(delta.Caption))) / 2 * Val(tx2a.Text))
lx23.Visible = True
lx24.Visible = True
    lx24.Left = 880
lx25.Visible = True
lx26.Visible = True
    lx26.Left = 720
rx21.Left = 1120
Else
rx21.Caption = "´Ë·½³ÌÎÞ½â"
rx22.Caption = ""
lx23.Visible = False
lx24.Visible = False
lx25.Visible = False
lx26.Visible = False

End If
delta.Caption = "¦¤="
delta.Caption = delta.Caption + Str((Val(tx2b.Text)) ^ 2 - 4 * Val(tx2a.Text) * Val(tx2c.Text))
End Sub

Private Sub tx2a_Click()
tx2a.Text = ""
End Sub


Private Sub tx2b_Click()
tx2b.Text = ""
End Sub


Private Sub tx2c_Click()
tx2c.Text = ""
End Sub


