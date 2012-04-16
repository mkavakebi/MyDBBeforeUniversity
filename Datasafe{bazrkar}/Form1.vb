Option Strict Off
Option Explicit On
Friend Class Form1

    Inherits System.Windows.Forms.Form

    Dim m As New Bitmap("D:\Documents and Settings\Mohammad\Desktop\New Folder\N.bmp")
    Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
        Dim i As Object
        Text1.Text = Mid(Text1.Text, 1, (Len(Text1.Text) \ 3) * 3)
        For i = 1 To Len(Text1.Text) Step 3
            m.SetPixel(i \ 3, 10, System.Drawing.ColorTranslator.FromOle(RGB(Asc(Mid(Text1.Text, i, 1)), Asc(Mid(Text1.Text, i + 1, 1)), Asc(Mid(Text1.Text, i + 2, 1)))))
        Next i
        Picture1.BackgroundImage = m
    End Sub

    Private Sub Command2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Command2.Click
        Dim i As Int32
        For i = 1 To 5
            Text2.Text += Chr(m.GetPixel(i, 10).R) + Chr(m.GetPixel(i, 10).G) + Chr(m.GetPixel(i, 10).B)
        Next
    End Sub
End Class