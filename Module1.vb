Private Sub CommandButton1_Click()
'Заполнение массива'
For i = 1 To 30
Cells(1, i) = Int((2000 * Rnd) - 1000)
Next i
End Sub

Private Sub CommandButton2_Click()
'Рассчитывает и выводит среднее арифметическое из элементов оканчивающихся на 5'
k = 0
a = 0
For i = 1 To 30
If Cells(1, i) Mod 10 = 5 Then
a = a + Cells(1, i)
k = k + 1
ElseIf Cells(1, i) Mod 10 = -5 Then
a = a + Cells(1, i)
k = k + 1
End If
Next i
b = a / k
MsgBox (b)
End Sub

Private Sub CommandButton3_Click()
'Закрытие формы'
UserForm1.Hide
End Sub