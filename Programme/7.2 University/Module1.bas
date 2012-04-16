Attribute VB_Name = "Module1"
Public window_style As String
Public Sub unload_load(ByVal Form As Form)
Unload Form
Load Form
Form.Visible = True
End Sub

Public Function coding(ByVal pass As String) As String
Randomize Timer
For i = 1 To Len(pass)
tic = Fix(Rnd * 200) + 10
r = r + Chr(tic)
f = Mid(pass, i, 1)
cit = tic + Asc(f)
If cit > 255 Then cit = cit - 255
t = t + Chr(cit)
Next
coding = r + t
End Function

Public Function decoding(ByVal pass As String) As String
k = Len(pass) / 2
For i = 1 To k
t1 = Mid(pass, i, 1)
t2 = Mid(pass, i + k, 1)
tic = Asc(t2) - Asc(t1)
If tic < 0 Then tic = tic + 255
r = r + Chr(tic)
Next
decoding = r
End Function

Public Sub comment_nexttime(ByVal comment As String)
free = FreeFile
Open App.Path + "\next_time.comment" For Input As #free
While EOF(free) = False
Input #free, co
If co <> "" Then
com = com + vbCrLf + co
End If
Wend
Close #free
'''''''''''''''''''''
Open App.Path + "\next_time.comment" For Output As #free
Print #free, com + vbCrLf + comment
Close #free
'''''''''''''''''''''
End Sub

Public Sub comment_now(ByVal comment As String)
Dim a As New show_comments_form
a.Show
a.Label1.Caption = comment
End Sub

Public Sub comment_dynamic(ByVal comment As String, ByVal code As String)
comment_now (comment)
free = FreeFile
Open App.Path + "\dynamic.comment" For Input As #free
While EOF(free) = False
Input #free, co
If co <> "" Then
com = com + vbCrLf + co
End If
Wend
Close #free
'''''''''''''''''''''
Open App.Path + "\dynamic.comment" For Output As #free
Print #free, com + vbCrLf + code + "=" + comment
Close #free
'''''''''''''''''''''
End Sub

Public Sub comment_dynamic_delet(ByVal code As String, Optional comment As String)
If comment <> "" Then
comment_now (comment)
End If
'''''''''''''''''''''''
free = FreeFile
Open App.Path + "\dynamic.comment" For Input As #free
While EOF(free) = False
Input #free, co
If Len(co) > Len(code) Then
If Mid(co, 1, Len(code)) = code Then
co = ""
End If
End If
If co <> "" Then
com = com + vbCrLf + co
End If
Wend
Close #free
'''''''''''''''''''''
Open App.Path + "\dynamic.comment" For Output As #free
Print #free, com
Close #free
'''''''''''''''''''''
End Sub

Public Sub comment_date(ByVal comment As String, ByVal dates As String)
If dates = Date Then
comment_now (comment)
End If
'''''''''''''''
free = FreeFile
Open App.Path + "\date.comment" For Input As #free
While EOF(free) = False
Input #free, co
If IsNumeric(co) Then co = CStr(co)
If co <> "" Then
com = com + vbCrLf + co
End If
Wend
Close #free
'''''''''''''''''''''
Open App.Path + "\date.comment" For Output As #free
Print #free, com + vbCrLf + dates + "=" + comment
Close #free
'''''''''''''''''''''
End Sub

Public Sub comment_date_delet(ByVal code As String, Optional dates As String)
If comment <> "" Then
comment_now (comment)
End If
'''''''''''''''''''''''
free = FreeFile
Open App.Path + "\date.comment" For Input As #free
While EOF(free) = False
Input #free, co
If Len(co) > Len(dates) Then
If Mid(co, 1, Len(dates)) = dates Then
co = ""
End If
End If
If co <> "" Then
com = com + vbCrLf + co
End If
Wend
Close #free
'''''''''''''''''''''
Open App.Path + "\date.comment" For Output As #free
Print #free, com
Close #free
'''''''''''''''''''''
End Sub

Public Function date_menha(ByVal a As String, ByVal b As String) As Integer
On Error Resume Next
a_sal = Val(Mid(a, 1, 4))
b_sal = Val(Mid(b, 1, 4))
a_mah = Val(Mid(a, 6, 2))
b_mah = Val(Mid(b, 6, 2))
a_roz = Val(Mid(a, 9, 2))
b_roz = Val(Mid(b, 9, 2))
''''''''''''''''
D = D + b_roz - a_roz
D = D + b_mah * 30 - a_mah * 30
D = D + b_sal * 365 - a_sal * 365
''''''''''''''''
date_menha = D
End Function

Public Sub load_un(ByRef a As Form)
'Dim o As New a
Load o
Unload a
o.Show
End Sub

Public Function mainshow()
If window_style = "window" Or window_style = "" Then
main.Show
Else
Main_2.Show
End If
End Function

Public Function mainhide()
If window_style = "window" Or window_style = "" Then
main.Hide
Else
'Main_2.Hide
End If
End Function
