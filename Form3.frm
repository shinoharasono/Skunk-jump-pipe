VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14040
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   625
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   936
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2760
      Top             =   6120
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   7080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1800
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1080
      Top             =   3120
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5040
      Top             =   1200
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
   Begin VB.Label Label2 
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
   Begin VB.Image Image1 
      Height          =   645
      Left            =   960
      Picture         =   "Form3.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   1
      Left            =   6000
      Picture         =   "Form3.frx":0EB5
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   3120
      Picture         =   "Form3.frx":C4A7
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   2
      Left            =   9000
      Picture         =   "Form3.frx":17A99
      Stretch         =   -1  'True
      Top             =   6480
      Width           =   1125
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   3
      Left            =   12480
      Picture         =   "Form3.frx":2308B
      Stretch         =   -1  'True
      Top             =   5640
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   3
      Left            =   12000
      Picture         =   "Form3.frx":2E67D
      Stretch         =   -1  'True
      Top             =   -3240
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   2
      Left            =   9600
      Picture         =   "Form3.frx":395C2
      Stretch         =   -1  'True
      Top             =   -2040
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   1
      Left            =   6240
      Picture         =   "Form3.frx":44507
      Stretch         =   -1  'True
      Top             =   -2640
      Width           =   1125
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   0
      Left            =   3120
      Picture         =   "Form3.frx":4F44C
      Stretch         =   -1  'True
      Top             =   -2280
      Width           =   1125
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z As Integer
Dim x As Integer
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
    
End Sub
Sub stopgame()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    Timer4.Enabled = False
    Timer5.Enabled = False
    
    Form7.Show
    
End Sub
Private Sub Form_Load()
    z = Image2(3).Left
    x = Image3(3).Left
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
        If Image2(a).Left + Image2(a).Width < 0 Then
            Image2(a).Left = z
            Image2(a).Tag = ""
        End If
        If Image3(a).Left + Image3(a).Width < 0 Then
            Image3(a).Left = x
            Image3(a).Tag = ""
        End If
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
