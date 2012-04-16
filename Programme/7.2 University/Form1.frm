VERSION 5.00
Begin VB.Form main 
   AutoRedraw      =   -1  'True
   Caption         =   "{»Â ‰«„ Œœ«}"
   ClientHeight    =   7455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   7455
   ScaleWidth      =   10860
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Õ÷Ê— Ê €Ì«» œ«‰‘ ¬„Ê“«‰"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1635
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4215
      Width           =   1920
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ò«—‰«„Â ê—ÊÂÌ "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1785
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5970
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ò«—‰«„Â ›—œÌ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1785
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5430
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "„⁄·„ ÃœÌœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1770
      Width           =   1920
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "À»  ‰«„"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1650
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   465
      Width           =   1920
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   3030
      Top             =   1950
   End
   Begin VB.Data data_date 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   255
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "security_time_backups"
      Top             =   270
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3870
      Top             =   165
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ê—ÊÂ ÃœÌœ"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1620
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2895
      Width           =   1920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "side bar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   9600
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   795
      Width           =   780
   End
   Begin VB.Image Image12 
      Height          =   1080
      Left            =   9495
      Picture         =   "Form1.frx":12ED2
      Top             =   240
      Width           =   1080
   End
   Begin VB.Image Image11 
      Height          =   1080
      Left            =   495
      Picture         =   "Form1.frx":13BEA
      Top             =   3975
      Width           =   1080
   End
   Begin VB.Image Image3 
      Height          =   1080
      Left            =   540
      Picture         =   "Form1.frx":14A2D
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Image Image2 
      Height          =   1080
      Left            =   3450
      Picture         =   "Form1.frx":1575A
      Top             =   5430
      Width           =   1080
   End
   Begin VB.Image Image10 
      Height          =   1080
      Left            =   525
      Picture         =   "Form1.frx":16503
      Top             =   2595
      Width           =   1080
   End
   Begin VB.Image Image8 
      Height          =   1080
      Left            =   3600
      Picture         =   "Form1.frx":174FA
      Top             =   225
      Width           =   1080
   End
   Begin VB.Image Image9 
      Height          =   1080
      Left            =   540
      Picture         =   "Form1.frx":18096
      Top             =   225
      Width           =   1080
   End
   Begin VB.Image Image7 
      Height          =   1080
      Left            =   3630
      Picture         =   "Form1.frx":18D0B
      Top             =   3975
      Width           =   1080
   End
   Begin VB.Image Image6 
      Height          =   1080
      Left            =   3555
      Picture         =   "Form1.frx":19B34
      Top             =   2595
      Width           =   1080
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«„‰Ì  Ê Å‘ Ì»«‰Ì"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   7455
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3120
      Width           =   1590
   End
   Begin VB.Image Image4 
      Height          =   1080
      Left            =   3645
      Picture         =   "Form1.frx":1A998
      Top             =   1365
      Width           =   1080
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "„— ÷Ì "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7890
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   6945
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "òÊ«ò»Ì"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6705
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   6945
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   10230
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   7065
      Width           =   495
   End
   Begin VB.Image Image5 
      Height          =   1080
      Left            =   480
      Picture         =   "Form1.frx":1B5D2
      Top             =   1365
      Width           =   1080
   End
   Begin VB.Image Image15 
      Appearance      =   0  'Flat
      Height          =   720
      Left            =   10080
      Picture         =   "Form1.frx":1C247
      Top             =   6780
      Width           =   720
   End
   Begin VB.Image Image1 
      Height          =   7965
      Left            =   -390
      Picture         =   "Form1.frx":1C9EA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   11925
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command1_Click()
Me.Hide
add_student.Show
End Sub

Private Sub Command2_Click()
Me.Hide
karname_form.Show
End Sub

Private Sub Command3_Click()
'comment_now (".«“ «” ›«œÂ Ì ‘„« «“ «Ì‰ »—‰«„Â Œ—”‰œÌ„")
Timer1.Enabled = True
SaveSetting "mk", "mk", "mk", "I'm close"
End Sub

Private Sub Command4_Click()
Me.Hide
add_groupform.Show
End Sub


Private Sub Command5_Click()
Karname2.Show
Me.Hide
End Sub

Private Sub Command6_Click()
add_teacher.Show
Me.Hide
End Sub



Private Sub Command9_Click()
Me.Hide
studentspresentation_form.Show
End Sub

Private Sub Form_Load()
'Image1.Move 0, 0
'Me.Width = Image1.Width
'Me.Height = Image1.Height
data_date.DatabaseName = App.Path + "\db1.mdb"
'Command1.Picture = ImageList1.ListImages(1).ExtractIcon
'Command9.Picture = ImageList1.ListImages(3).ExtractIcon
'Command3.Picture = ImageList1.ListImages(2).ExtractIcon
'comment_now ("Ê—Êœ ‘„« —« »Â »—‰«„Â ŒÌ— „ﬁœ„ ⁄—÷ „Ì ‰„«ÌÌ„.")
End Sub

Private Sub Form_Unload(Cancel As Integer)
'comment_now (".«“ «” ›«œÂ Ì ‘„« «“ «Ì‰ »—‰«„Â ŒÊ‘ Œ—”‰œÌ„")
Timer1.Enabled = True
SaveSetting "mk", "mk", "mk", "I'm close"
End
End Sub

Private Sub Image12_Click()
Me.Hide
window_style = "sidebar"
SaveSetting "mk", "mk", "style", "sidebar"
Main_2.Show
End Sub

Private Sub Image15_Click()
 Label1_Click
End Sub

Private Sub Label1_Click()
Me.WindowState = vbMinimized
'comment_now (".«“ «” ›«œÂ Ì ‘„« «“ «Ì‰ »—‰«„Â Œ—”‰œÌ„")
Timer1.Enabled = True
SaveSetting "mk", "mk", "mk", "I'm close"
End Sub

Private Sub Label4_Click()
Image12_Click
End Sub

Private Sub Label7_Click()
Me.Hide
backups_security_form.Show
End Sub

Private Sub Timer1_Timer()
End
End Sub

Private Sub Timer2_Timer()
Static c As Integer
c = c + 1
If c < 100 Then Exit Sub
Label2.Top = Label2.Top - 10
Label3.Top = Label3.Top - 10
If c > 170 Then Timer2.Enabled = False

End Sub
