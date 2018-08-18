VERSION 5.00
Begin VB.Form wx3 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ò»ÔªÈý´Î·½³Ì"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton sx3 
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
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1560
      Width           =   1096
   End
   Begin VB.TextBox tx3d 
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
      Left            =   5040
      TabIndex        =   3
      Text            =   "d"
      Top             =   240
      Width           =   1092
   End
   Begin VB.TextBox tx3a 
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
   Begin VB.TextBox tx3b 
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
   Begin VB.TextBox tx3c 
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
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1096
   End
   Begin VB.Label lx3n 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3720
      TabIndex        =   12
      Top             =   3360
      Width           =   2175
   End
   Begin VB.Label lx3m 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label lx3v 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lx3u 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label lx33 
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
      Height          =   195
      Left            =   3200
      TabIndex        =   8
      Top             =   240
      Width           =   105
   End
   Begin VB.Label lx32 
      BackColor       =   &H00FFFFFF&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   6.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Left            =   1560
      TabIndex        =   6
      Top             =   240
      Width           =   100
   End
   Begin VB.Label lx31 
      BackColor       =   &H00FFFFFF&
      Caption         =   "x  +                     x  +                    x +          +        =0 "
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
      TabIndex        =   7
      Top             =   270
      Width           =   5295
   End
   Begin VB.Menu mAbout 
      Caption         =   "¹ØÓÚ"
   End
End
Attribute VB_Name = "wx3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lx3v_Click()
lx3u.Caption = Str((9 * Val(tx3a.Text) * Val(tx3b.Text) * Val(tx3c.Text) - 27 * (Val(tx3a.Text)) ^ 2 * Val(tx3d.Text) - 2 * (Val(tx3b.Text)) ^ 3) / 54 * (Val(tx3a.Text)) ^ 3)
lx3v.Caption = Sqr(12 * Val(tx3a.Text) * (Val(tx3b.Text)) ^ 3 - (Val(tx3b.Text)) ^ 2 * (Val(tx3c.Text)) ^ 2 - 54 * Val(tx3a.Text) * Val(tx3b.Text) * Val(tx3c.Text) * Val(tx3d.Text) + 81 * (Val(tx3a.Text)) ^ 2 * (Val(tx3d.Text)) ^ 2 + 12 * (Val(tx3b.Text)) ^ 3 * Val(tx3d.Text)) / (18 * (Val(tx3a.Text)) ^ 2)

If Abs(Val(lx3u.Caption) + Val(lx3v.Caption)) >= Abs(Val(lx3u.Caption) - Val(lx3v.Caption)) Then
lx3m.Caption = (Val(lx3u.Caption) + Val(lx3v.Caption)) ^ (1 / 3)
Else
lx3m.Caption = (Val(lx3u.Caption) - Val(lx3v.Caption)) ^ (1 / 3)
End If
End Sub

Private Sub mAbout_Click()
wAbout.Show 1
End Sub


Private Sub mBack_Click()
Unload Me
End Sub


Private Sub sx2_Click()

End Sub


Private Sub tx2a_Change()

End Sub


Private Sub tx2a_Click()

End Sub


Private Sub sx3_Click()
lx3u.Caption = Str((9 * Val(tx3a.Text) * Val(tx3b.Text) * Val(tx3c.Text) - 27 * (Val(tx3a.Text)) ^ 2 * Val(tx3d.Text) - 2 * (Val(tx3b.Text)) ^ 3) / 54 * (Val(tx3a.Text)) ^ 3)
lx3v.Caption = Sqr(12 * Val(tx3a.Text) * (Val(tx3b.Text)) ^ 3 - (Val(tx3b.Text)) ^ 2 * (Val(tx3c.Text)) ^ 2 - 54 * Val(tx3a.Text) * Val(tx3b.Text) * Val(tx3c.Text) * Val(tx3d.Text) + 81 * (Val(tx3a.Text)) ^ 2 * (Val(tx3d.Text)) ^ 2 + 12 * (Val(tx3b.Text)) ^ 3 * Val(tx3d.Text)) / (18 * (Val(tx3a.Text)) ^ 2)




End Sub
Private Sub tx3a_Click()
tx3a.Text = ""
End Sub


Private Sub tx3b_Click()
tx3b.Text = ""
End Sub


Private Sub tx3c_Click()
tx3c.Text = ""
End Sub


Private Sub tx3d_Click()
tx3d.Text = ""
End Sub


