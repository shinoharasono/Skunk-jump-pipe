VERSION 5.00
Begin VB.Form Form5 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form5"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   21090
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   626
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1406
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer7 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6960
      Top             =   5760
   End
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4440
      Top             =   6360
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   6600
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1080
      Top             =   6120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1440
      Top             =   2280
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   600
      Top             =   2520
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2640
      Top             =   2400
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
      TabIndex        =   1
      Top             =   0
      Width           =   1095
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
      TabIndex        =   0
      Top             =   0
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   3
      Left            =   20760
      Picture         =   "Form5.frx":0000
      Stretch         =   -1  'True
      Top             =   -4920
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   2
      Left            =   8880
      Picture         =   "Form5.frx":AF45
      Stretch         =   -1  'True
      Top             =   -6000
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   1
      Left            =   6120
      Picture         =   "Form5.frx":15E8A
      Stretch         =   -1  'True
      Top             =   -3840
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   0
      Left            =   3240
      Picture         =   "Form5.frx":20DCF
      Stretch         =   -1  'True
      Top             =   -2520
      Width           =   1125
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   480
      Picture         =   "Form5.frx":2BD14
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   1
      Left            =   6720
      Picture         =   "Form5.frx":2CBC9
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   2040
      Picture         =   "Form5.frx":381BB
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   2
      Left            =   9480
      Picture         =   "Form5.frx":437AD
      Stretch         =   -1  'True
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   3
      Left            =   20880
      Picture         =   "Form5.frx":4ED9F
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1125
   End
End
Attribute VB_Name = "Form5"
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
    Timer7.Enabled = True
End Sub
Sub stopgame()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    Timer7.Enabled = False
    Form15.Show
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
        Image2(a).Left = Image2(a).Left - 20
        Image3(a).Left = Image3(a).Left - 20
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
        Timer7.Enabled = False
        Form14.Show
        
    End If
End Sub


Private Sub Timer7_Timer()
    Dim a As Integer
    For a = 0 To 3
        If Image1.Left > Image3(2).Left + Image3(2).Width Then
        Image2(a).Left = Image2(a).Left - 35
        Image3(a).Left = Image3(a).Left - 35
        End If
    Next a
End Sub
