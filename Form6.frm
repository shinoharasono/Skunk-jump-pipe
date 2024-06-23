VERSION 5.00
Begin VB.Form Form6 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form6"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15540
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   642
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1036
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   10440
      Top             =   5880
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7200
      Top             =   4560
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6120
      Top             =   6000
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2880
      Top             =   1560
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   1800
      Top             =   1440
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5160
      Top             =   1200
   End
   Begin VB.Image Image2 
      Height          =   6765
      Index           =   0
      Left            =   8520
      Picture         =   "Form6.frx":0000
      Stretch         =   -1  'True
      Top             =   -3840
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "按空格臭鼬就会飞"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin VB.Image Image1 
      Height          =   645
      Left            =   2280
      Picture         =   "Form6.frx":AF45
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   720
   End
   Begin VB.Image Image3 
      Height          =   6735
      Index           =   0
      Left            =   8640
      Picture         =   "Form6.frx":BDFA
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1125
   End
End
Attribute VB_Name = "Form6"
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
    Form11.Show
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
    For a = 0 To 0
        Image2(a).Left = Image2(a).Left - 15
        Image3(a).Left = Image3(a).Left - 15
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
    For d = 0 To 0
        If PZ(Image1, Image2(d)) Or PZ(Image1, Image3(d)) Then
            Call stopgame
        End If
    Next d
End Sub


Private Sub Timer6_Timer()
    If Image1.Left > Image3(0).Left + Image3(0).Width Then
        Timer1.Enabled = False
        Timer2.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        Timer5.Enabled = False
        Form10.Show
        
    End If
End Sub


