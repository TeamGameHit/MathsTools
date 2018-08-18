VERSION 5.00
Begin VB.Form wfx1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Ò»´Îº¯Êý"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   690
   ClientWidth     =   2745
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   2745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ËùÓÐÕßÖÐÐÄ
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
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2640
      Width           =   1096
   End
   Begin VB.CommandButton sfx1 
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
      TabIndex        =   9
      Top             =   2640
      Width           =   1096
   End
   Begin VB.TextBox tfx1y2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   5
      Text            =   "y"
      Top             =   1440
      Width           =   1200
   End
   Begin VB.TextBox tfx1x2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Text            =   "x"
      Top             =   1440
      Width           =   1200
   End
   Begin VB.TextBox tfx1y1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Text            =   "y"
      Top             =   600
      Width           =   1200
   End
   Begin VB.TextBox tfx1x1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Text            =   "x"
      Top             =   600
      Width           =   1200
   End
   Begin VB.Label rfx1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "f(x)="
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label lfx11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ò»´Îº¯Êý y=f(x) ¹ýµã"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label lfx13 
      BackColor       =   &H00FFFFFF&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1320
      TabIndex        =   6
      Top             =   1440
      Width           =   135
   End
   Begin VB.Label lfx12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Óë"
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   3
      Top             =   960
      Width           =   375
   End
   Begin VB.Label lfx12 
      BackColor       =   &H00FFFFFF&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Î¢ÈíÑÅºÚ"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   2
      Top             =   600
      Width           =   135
   End
   Begin VB.Menu mAbout 
      Caption         =   "¹ØÓÚ"
   End
End
Attribute VB_Name = "wfx1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mAbout_Click()
wAbout.Show 1
End Sub

Private Sub mBack_Click()
Unload Me
End Sub


Private Sub sx1_Click()

End Sub


Private Sub sfx1_Click()
rfx1.Caption = "f(x)="
rfx1.Caption = rfx1.Caption + Str((Val(tfx1y2) - Val(tfx1y1)) / (Val(tfx1x2) - Val(tfx1x1))) + "x+" + Str(Val(tfx1y2) - Val(tfx1x2) * (Val(tfx1y2) - Val(tfx1y1)) / (Val(tfx1x2) - Val(tfx1x1)))
End Sub

Private Sub tfx1x1_Click()
tfx1x1.Text = ""
End Sub
Private Sub tfx1x2_Click()
tfx1x2.Text = ""
End Sub
Private Sub tfx1y1_Click()
tfx1y1.Text = ""
End Sub
Private Sub tfx1y2_Click()
tfx1y2.Text = ""
End Sub
