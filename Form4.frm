VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form4"
   ClientHeight    =   9180
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   20850
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   612
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4320
      Top             =   7080
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   6360
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8520
      Top             =   4920
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   840
      Top             =   2160
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4200
      Top             =   1920
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "得分："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   5
      Left            =   19680
      Picture         =   "Form4.frx":0000
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   4
      Left            =   16440
      Picture         =   "Form4.frx":B5F2
      Stretch         =   -1  'True
      Top             =   6360
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   5
      Left            =   19560
      Picture         =   "Form4.frx":16BE4
      Stretch         =   -1  'True
      Top             =   -2880
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   4
      Left            =   16920
      Picture         =   "Form4.frx":21B29
      Stretch         =   -1  'True
      Top             =   -1920
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   3
      Left            =   13680
      Picture         =   "Form4.frx":2CA6E
      Stretch         =   -1  'True
      Top             =   -2760
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   2
      Left            =   10320
      Picture         =   "Form4.frx":379B3
      Stretch         =   -1  'True
      Top             =   -3480
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   1
      Left            =   7320
      Picture         =   "Form4.frx":428F8
      Stretch         =   -1  'True
      Top             =   -3240
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   0
      Left            =   3960
      Picture         =   "Form4.frx":4D83D
      Stretch         =   -1  'True
      Top             =   -2640
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1320
      Picture         =   "Form4.frx":58782
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   1
      Left            =   7320
      Picture         =   "Form4.frx":59637
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   3840
      Picture         =   "Form4.frx":64C29
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   2
      Left            =   10200
      Picture         =   "Form4.frx":7021B
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   3
      Left            =   13200
      Picture         =   "Form4.frx":7B80D
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   1125
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call opergame
        Image1.Top = Image1.Top - 50
    End If
End Sub

Sub opergame()
    Timer1.Enabled = True
    Timer2.Enabled = True
    Timer3.Enabled = True
    Timer4.Enabled = True
    Timer5.Enabled = True
    Timer6.Enabled = True
End Sub
Sub stopgame()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Form13.Show
End Sub


Private Sub Timer1_Timer()
    Image1.Top = Image1.Top + 5
End Sub

Private Sub Timer2_Timer()
    If Image1.Left > 168 Then
        Timer2.Enabled = False
    End If
    Image1.Left = Image1.Left + 1
    
End Sub

Private Sub Timer3_Timer()
    Dim a As Integer
    For a = 0 To 5
        Image2(a).Left = Image2(a).Left - 15
        Image3(a).Left = Image3(a).Left - 15
    Next a
End Sub

Private Sub Timer4_Timer()
    Dim a As Integer
    For a = 0 To 5
        
        If Image2(a).Tag = "" And Image2(a).Left + Image2(a).Width < Image1.Left Then
            Label3.Caption = Val(Label3.Caption) + 1
            Image2(a).Tag = "1"
        End If
        If Image3(a).Tag = "" And Image3(a).Left + Image3(a).Width < Image1.Left Then
            Label3.Caption = Val(Label3.Caption) + 1
            Image3(a).Tag = "1"
        End If
    Next a
End Sub
Function PZ(E As Image, B As Image)
    Dim c As Boolean
        c = False
          If E.Left + E.Width > B.Left And E.Left < B.Left + B.Width And E.Top + E.Height > B.Top And E.Top < B.Top + B.Height Then
            c = True
        End If
        PZ = c
End Function

Private Sub Timer5_Timer()
    Dim d As Integer
    For d = 0 To 5
        If PZ(Image1, Image2(d)) Or PZ(Image1, Image3(d)) Then
            Call stopgame
        End If
    Next d
End Sub


Private Sub Timer6_Timer()
    If Image1.Left > Image3(5).Left + Image3(5).Width Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Form12.Show
        
    End If
End Sub


