VERSION 5.00
Begin VB.Form Main_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "       SideBar"
   ClientHeight    =   4395
   ClientLeft      =   5160
   ClientTop       =   3945
   ClientWidth     =   2205
   Icon            =   "Main_2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   2205
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   465
      Top             =   195
   End
   Begin VB.Data data_date 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   900
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "security_time_backups"
      Top             =   165
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   0
      Picture         =   "Main_2.frx":12ED2
      ToolTipText     =   "ﬂ«—‰«„Â ê—ÊÂÌ"
      Top             =   2190
      Width           =   1110
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   1095
      Picture         =   "Main_2.frx":13BFF
      ToolTipText     =   "«÷«›Â ”«“Ì „⁄·„"
      Top             =   1095
      Width           =   1110
   End
   Begin VB.Image Image6 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   1095
      Picture         =   "Main_2.frx":14839
      ToolTipText     =   "«÷«›Â ”«“Ì ê—ÊÂ"
      Top             =   2190
      Width           =   1110
   End
   Begin VB.Image Image7 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   1095
      Picture         =   "Main_2.frx":1569D
      ToolTipText     =   "Õ÷Ê— Ê €Ì«» œ«‰‘ ¬„Ê“«‰"
      Top             =   3285
      Width           =   1110
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   0
      Picture         =   "Main_2.frx":164C6
      ToolTipText     =   "ﬂ«—‰«„Â ›—œÌ"
      Top             =   1095
      Width           =   1110
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   0
      Picture         =   "Main_2.frx":1726F
      ToolTipText     =   "«„‰Ì  Ê Å‘ Ì»«‰Ì"
      Top             =   3285
      Width           =   1110
   End
   Begin VB.Image Image8 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   1095
      Picture         =   "Main_2.frx":17E6C
      ToolTipText     =   "À»  ‰«„"
      Top             =   0
      Width           =   1110
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Window"
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   570
      Width           =   825
   End
   Begin VB.Image Image12 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1110
      Left            =   0
      Picture         =   "Main_2.frx":18A08
      ToolTipText     =   " »œÌ· »Â Å‰Ã—Â"
      Top             =   0
      Width           =   1110
   End
End
Attribute VB_Name = "Main_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Command3_Click()
'comment_now (".«“ «” ›«œÂ Ì ‘„« «“ «Ì‰ »—‰«„Â Œ—”‰œÌ„")
Timer1.Enabled = True
SaveSetting "mk", "mk", "mk", "I'm close"
End Sub

Private Sub Form_Activate()
Form_Resize
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

Private Sub Form_Resize()
Me.Left = Screen.Width - Me.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
'comment_now (".«“ «” ›«œÂ Ì ‘„« «“ «Ì‰ »—‰«„Â ŒÊ‘ Œ—”‰œÌ„")
Timer1.Enabled = True
SaveSetting "mk", "mk", "mk", "I'm close"
End
End Sub

Private Sub Image1_Click()
'Me.Hide
backups_security_form.Show
End Sub

Private Sub Image12_Click()
Me.Hide
window_style = "window"
SaveSetting "mk", "mk", "style", "window"
main.Show
End Sub

Private Sub Image2_Click()
'Me.Hide
karname_form.Show
End Sub

Private Sub Image3_Click()
Karname2.Show
'Me.Hide
End Sub

Private Sub Image4_Click()
add_teacher.Show
'Me.Hide
End Sub

Private Sub Image6_Click()
'Me.Hide
add_groupform.Show
End Sub

Private Sub Image7_Click()
'Me.Hide
studentspresentation_form.Show
End Sub

Private Sub Image8_Click()
'Me.Hide
add_student.Show
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

Private Sub Timer1_Timer()
End
End Sub
