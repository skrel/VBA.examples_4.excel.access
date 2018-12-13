Attribute VB_Name = "Module3"
Sub WriteFile()

ThisFile = ThisWorkbook.Path & Application.PathSeparator & "Results.txt"

Open ThisFile For Append As #1
FinalRow = Range("G1").End(xlUp).Row

For j = 1 To FinalRow
Print #1, Cells(j, 1).Value
Next j
Close #1
MsgBox ThisFile & " completed."
End Sub

