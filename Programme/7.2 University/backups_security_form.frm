VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form backups_security_form 
   Caption         =   "«„‰Ì  Ê Å‘ Ì»«‰Ì"
   ClientHeight    =   3435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4575
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
   RightToLeft     =   -1  'True
   ScaleHeight     =   3435
   ScaleWidth      =   4575
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1530
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4335
      Begin VB.CommandButton Command3 
         Caption         =   " «ÌÌœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   90
         Picture         =   "backups_security_form.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   330
         Width           =   840
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "$"
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "!"
         TabIndex        =   3
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         PasswordChar    =   "$"
         TabIndex        =   4
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label7 
         Caption         =   ":  ò—«—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   ": —„“ ⁄»Ê— ÃœÌœ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   ": —„“ ⁄»Ê— ﬁ»·Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   13
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "—„“ ⁄»Ê—"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   12
         Top             =   0
         Width           =   780
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4335
      Begin VB.CommandButton Command1 
         Caption         =   "”«Œ ‰ ›«Ì· Å‘ Ì»«‰"
         Height          =   705
         Left            =   2175
         Picture         =   "backups_security_form.frx":0504
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   795
         Width           =   1695
      End
      Begin VB.CommandButton Command5 
         Caption         =   "«” ›«œÂ «“ ›«Ì· Å‘ Ì»«‰"
         Height          =   705
         Left            =   480
         Picture         =   "backups_security_form.frx":0A58
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   795
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         Caption         =   "$"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   255
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   720
         TabIndex        =   0
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "›«Ì· Å‘ Ì»«‰Ì"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   10
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   ": ¬œ—” ›«Ì· Å‘ Ì»«‰"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   1455
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   " ·ÿ›« ›«Ì· ŒÊœ —« «‰ Œ«Ì ò‰»œ"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontName        =   "arial"
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "security_time_backups"
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "backups_security_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a As New FileSystemObject
If Right(Text2.Text, 4) <> ".mdb" Then
ff = Text2.Text
ff = ff + ".mdb"
a.CopyFile App.Path + "\db1.mdb", ff
Else
a.CopyFile App.Path + "\db1.mdb", Text2.Text
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
CommonDialog1.ShowSave
If CommonDialog1.FileName <> "" Then
Text2.Text = CommonDialog1.FileName
End If
End Sub

Private Sub Command3_Click()
'''''''''''''''''''''''''''''
If data_date.Recordset.RecordCount = 0 Then
data_date.Recordset.AddNew
data_date.Recordset.Update
End If
data_date.Recordset.MoveFirst
pass = data_date.Recordset.Fields("entrance_password")
If IsNull(pass) Then pass = ""
If decoding(pass) <> LCase(text4.Text) Then
MsgBox "—„“ ⁄»Ê— ›⁄·Ì ‘„« ’ÕÌÕ ‰Ì” .", vbInformation
text4.Text = ""
text4.SetFocus
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''
If LCase(Text3.Text) <> LCase(text5.Text) Then
MsgBox "—„“ ⁄»Ê— —« Ìò»«— œÌê— Ê«—œ ò‰Ìœ!", vbCritical
Text3.Text = ""
text5.Text = ""
Exit Sub
End If
'''''''''''''''''''''''''''''''''''''''''
data_date.Recordset.Edit
data_date.Recordset.Fields("entrance_password") = coding(LCase(Text3.Text))
data_date.Recordset.Update
comment_now ("—„“ ⁄»Ê— »« „Ê›ﬁÌ  À»  ‘œ")
End Sub

Private Sub Command5_Click()
Dim a As New FileSystemObject
On Error Resume Next
CommonDialog1.ShowOpen
If CommonDialog1.FileName <> "" Then
a.CopyFile CommonDialog1.FileName, App.Path + "\db1.mdb"
End If
End Sub

Private Sub Form_Load()
Text2.Text = App.Path
data_date.DatabaseName = App.Path + "\db1.mdb"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Me.Hide
mainshow
End Sub
