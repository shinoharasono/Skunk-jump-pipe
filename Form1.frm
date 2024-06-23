VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "臭鼬"
   ClientHeight    =   9405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16020
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   627
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1068
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "无限关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   8640
      TabIndex        =   4
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "加速关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   4080
      TabIndex        =   3
      Top             =   7680
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "第三关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   10200
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "第二关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   6360
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一关"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2520
      TabIndex        =   0
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   4485
      Left            =   5760
      Picture         =   "Form1.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3720
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
  Private Sub Command1_Click()
    Form1.Hide
    Form6.Show
End Sub

Private Sub Command2_Click()
    Form2.Show
    Form1.Hide
End Sub
Private Sub Command3_Click()
    Form4.Show
    Form1.Hide
End Sub
Private Sub Command4_Click()
    Form5.Show
    Form1.Hide
End Sub
Private Sub Command5_Click()
    Form3.Show
    Form1.Hide
End Sub


