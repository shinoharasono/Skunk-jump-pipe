VERSION 5.00
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form10"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "下一关"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "主界面"
      Height          =   615
      Left            =   840
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "您已经让会爬的臭鼬飞了，请前往下一关吧！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3375
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Form6
    Form10.Hide
    Form1.Show
End Sub

Private Sub Command2_Click()
    Unload Form6
    Form10.Hide
    Form2.Show
End Sub

