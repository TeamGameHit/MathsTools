VERSION 5.00
Begin VB.Form wAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "关于"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3465
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   10.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton mConfirm 
      BackColor       =   &H00FFFFFF&
      Caption         =   "确认"
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
      Left            =   795
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   345
      Picture         =   "wAbout.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   840
      Width           =   2895
   End
   Begin VB.Label mSM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "数学工具 by bac0id"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   225
      TabIndex        =   2
      Top             =   240
      Width           =   2835
   End
End
Attribute VB_Name = "wAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub mConfirm_Click()
Unload Me
End Sub


