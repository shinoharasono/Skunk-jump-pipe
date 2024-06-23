VERSION 5.00
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form7"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "主界面"
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "重来"
      Height          =   615
      Left            =   2520
      TabIndex        =   1
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "游戏结束！"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
    Unload Form3
    Form3.Show
End Sub

Private Sub Command2_Click()
    Form7.Hide
    Form3.Hide
    Form1.Show
End Sub

