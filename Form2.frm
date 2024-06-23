VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   9315
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15510
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   621
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1034
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   9960
      Top             =   4800
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4200
      Top             =   6240
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5760
      Top             =   6960
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4800
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1440
      Top             =   2640
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   2760
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
      Index           =   3
      Left            =   14520
      Picture         =   "Form2.frx":0000
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   2
      Left            =   10680
      Picture         =   "Form2.frx":B5F2
      Stretch         =   -1  'True
      Top             =   7800
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   4440
      Picture         =   "Form2.frx":16BE4
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   1
      Left            =   7560
      Picture         =   "Form2.frx":221D6
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   1560
      Picture         =   "Form2.frx":2D7C8
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   720
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   0
      Left            =   4440
      Picture         =   "Form2.frx":2E67D
      Stretch         =   -1  'True
      Top             =   -2760
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   1
      Left            =   7320
      Picture         =   "Form2.frx":395C2
      Stretch         =   -1  'True
      Top             =   -1920
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   2
      Left            =   10560
      Picture         =   "Form2.frx":44507
      Stretch         =   -1  'True
      Top             =   -2400
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   3
      Left            =   13920
      Picture         =   "Form2.frx":4F44C
      Stretch         =   -1  'True
      Top             =   -1200
      Width           =   1125
   End
End
Attribute VB_Name = "Form2"
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
    Form9.Show
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
    For a = 0 To 3
        Image2(a).Left = Image2(a).Left - 15
        Image3(a).Left = Image3(a).Left - 15
    Next a
End Sub

Private Sub Timer4_Timer()
    Dim a As Integer
    For a = 0 To 3
        
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
    For d = 0 To 3
        If PZ(Image1, Image2(d)) Or PZ(Image1, Image3(d)) Then
            Call stopgame
        End If
    Next d
End Sub


Private Sub Timer6_Timer()
    If Image1.Left > Image3(3).Left + Image3(3).Width Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Form8.Show
        
    End If
End Sub
