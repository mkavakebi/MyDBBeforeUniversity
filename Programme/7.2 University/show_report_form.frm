VERSION 5.00
Begin VB.Form show_report_form 
   Caption         =   "Form1"
   ClientHeight    =   450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   450
   ScaleWidth      =   1725
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
End
Attribute VB_Name = "show_report_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CrystalReport1.ReportFileName = App.Path + "\report5.rpt"
CrystalReport_kodakan.ReportFileName = App.Path + "\report6.rpt"
End Sub

Public Sub show_report()
Randomize Timer
CrystalReport1.RetrieveDataFiles
CrystalReport1.PrintReport
comment_now ("ò«—‰«„Â »« „Ê›ﬁÌ   ‰ŸÌ„ ê—œÌœ.")
End Sub

Public Sub show_report_Kodakan()
Randomize Timer
CrystalReport_kodakan.RetrieveDataFiles
CrystalReport_kodakan.PrintReport
comment_now ("ò«—‰«„Â »« „Ê›ﬁÌ   ‰ŸÌ„ ê—œÌœ.")
End Sub
