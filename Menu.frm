VERSION 5.00
Begin VB.Form Menu 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "菜单"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   495
   ClientWidth     =   6675
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Menu"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6675
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton mx2 
      Caption         =   "一元二次方程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   480
      TabIndex        =   1
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton mx3 
      Caption         =   "一元三次方程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton mx1 
      Caption         =   "一元一次方程"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   480
      TabIndex        =   0
      ToolTipText     =   "ax=b"
      Top             =   720
      Width           =   2535
   End
   Begin VB.Frame mfr2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "求函数解析式"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   15.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   3480
      TabIndex        =   7
      Top             =   120
      Width           =   3000
      Begin VB.CommandButton mfx3 
         Caption         =   "三次函数"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2535
      End
      Begin VB.CommandButton mfx2 
         Caption         =   "二次函数"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   4
         Top             =   1320
         Width           =   2535
      End
      Begin VB.CommandButton mfx1 
         Caption         =   "一次函数"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   2535
      End
   End
   Begin VB.Frame mfr1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "解方程"
      BeginProperty Font 
         Name            =   "微软雅黑 Light"
         Size            =   15.75
         Charset         =   134
         Weight          =   290
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   3000
   End
   Begin VB.Menu mAbout 
      Caption         =   "关于"
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
ml = "menu"
End Sub

Private Sub mAbout_Click()
wAbout.Show 1
End Sub

Private Sub mfx1_Click()
wfx1.Show
End Sub

Private Sub mx1_Click(Index As Integer)
wx1.Show
End Sub


Private Sub mx2_Click(Index As Integer)
wx2.Show
End Sub


Private Sub mx3_Click()
wx3.Show
End Sub


