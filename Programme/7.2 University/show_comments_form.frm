VERSION 5.00
Begin VB.Form show_comments_form 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   870
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   870
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2760
      Top             =   120
   End
   Begin VB.Timer Timer_down 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   840
      Top             =   120
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   120
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   375
      TabIndex        =   0
      Top             =   60
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   525
      Left            =   1260
      Picture         =   "show_comments_form.frx":0000
      Stretch         =   -1  'True
      Top             =   240
      Width           =   3285
   End
End
Attribute VB_Name = "show_comments_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Left = Screen.Width - Width - 500
Top = Screen.Height - 500
Image1.Move 0, 0, Width, Height
End Sub

Private Sub Timer_down_Timer()
Me.Top = Me.Top + 10
If Me.Top >= Screen.Height - 500 Then
Timer1.Enabled = False
Unload Me
End If
End Sub

Private Sub Timer1_Timer()
Me.Top = Me.Top - 10
If Me.Top <= Screen.Height - Height - 600 Then Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Static v As Integer
v = v + 1
If v = 6 Then
Timer2.Enabled = True
Timer_down.Enabled = True
End If
End Sub
