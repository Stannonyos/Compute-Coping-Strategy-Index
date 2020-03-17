Option Explicit
Private Sub compute_Click()
Dim i As Integer
For i = 2 To 12001
Cells(i, 5) = Cells(i, 3) * Cells(i, 4)
Next i
Range("n4").Value = Application.WorksheetFunction.AverageIfs(Range("e2:e12001"), Range("a2:a12001"), "=A")
Range("o4").Value = Application.WorksheetFunction.AverageIfs(Range("e2:e12001"), Range("a2:a12001"), "=B")
End Sub

Private Sub generate_Click()
Dim i As Integer
For i = 2 To 1001
Cells(i, 2).Value = Application.WorksheetFunction.RandBetween(1, 12)
Cells(i, 3).Value = Application.WorksheetFunction.RandBetween(1, 6)
Cells(i, 1).Value = Application.WorksheetFunction.RandBetween(1, 2)

''''''''''Counties''''''''''''''''''''''
If Cells(i, 1) = 1 Then
Cells(i, 1) = "A"
Else
Cells(i, 1) = "B"
End If

'''''''''''Strategies''''''''''''
If Cells(i, 2) = 1 Then
Cells(i, 2).Value = "Rely on less preferred and less expensive foods"   'Dietary Change'
Cells(i, 4).Value = 1
ElseIf Cells(i, 2).Value = 2 Then   'Increase Short-Term Household Food Availability'
Cells(i, 2).Value = "Borrow food from a friend or relative"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 3 Then
Cells(i, 2).Value = "Purchase food on credit"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 4 Then
Cells(i, 2).Value = "Gather wild food, hunt, or harvest immature crops"
Cells(i, 4) = 4
ElseIf Cells(i, 2).Value = 5 Then
Cells(i, 2).Value = "Consume seed stock held for next season"
Cells(i, 4) = 3

ElseIf Cells(i, 2).Value = 6 Then       'Decrease number of people'
Cells(i, 2).Value = "Send children to eat with neighbors"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 7 Then
Cells(i, 2).Value = "Send household members to beg"
Cells(i, 4) = 4
ElseIf Cells(i, 2).Value = 8 Then      'Rationing strategies'
Cells(i, 2).Value = "Limit portion size at mealtimes"
Cells(i, 4) = 1
ElseIf Cells(i, 2).Value = 9 Then
Cells(i, 2).Value = "Restrict consumption by adults in order for small children to eat"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 10 Then
Cells(i, 2).Value = "Feed working members of HH at the expense of non-working members"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 11 Then
Cells(i, 2).Value = "Reduce number of meals eaten in a day"
Cells(i, 4) = 2
ElseIf Cells(i, 2).Value = 12 Then
Cells(i, 2).Value = "Skip entire days without eating"
Cells(i, 4) = 4
End If
Next i
End Sub
Private Sub clear_Click()
Worksheets("Sheet1").Range("A2:f1000").clear
Worksheets("Sheet1").Range("n4").clear
Worksheets("Sheet1").Range("o4").clear
End Sub


