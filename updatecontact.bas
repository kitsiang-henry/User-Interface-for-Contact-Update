Attribute VB_Name = "Module1"
Sub Check_Update_Email()
    Dim Email As Range
    Set Email = Sheet1.Range("V2:V438")
    If WorksheetFunction.CountIf(Email, Sheet2.Range("C5")) > 0 Then
        MsgBox "The details of " & Sheet2.Range("C5") & " has been updated."
    Else
        MsgBox Sheet2.Range("C5") & " is not found as a primary email address."
        Exit Sub
    End If
    For Each Email In ActiveSheet.UsedRange
        If UCase(Email) = UCase(Sheet2.Range("C5")) Then
            Sheet3.Range("B2").Value = CStr(WorksheetFunction.Match(Sheet2.Range("C5"), Sheet1.Range("V1:V438"), 0))
            Sheet1.Range("AA" & Sheet3.Range("B2")) = Sheet2.Range("E5")
            Sheet3.Range("B3").Value = CStr(WorksheetFunction.Match(Sheet2.Range("C5"), Sheet1.Range("V1:V438"), 0))
            Sheet1.Range("AC" & Sheet3.Range("B3")) = Now
        End If
    Next Email
End Sub

Sub Check_Update_Email2()
    Dim Email2 As Range
    Set Email2 = Sheet1.Range("W2:W438")
    If WorksheetFunction.CountIf(Email2, Sheet2.Range("C5")) > 0 Then
        MsgBox "The details of " & Sheet2.Range("C5") & " has been updated."
    Else
        MsgBox Sheet2.Range("C5") & " is not found as a secondary email address."
        Exit Sub
    End If
    For Each Email2 In ActiveSheet.UsedRange
        If UCase(Email2) = UCase(Sheet2.Range("C5")) Then
            Sheet3.Range("B2").Value = CStr(WorksheetFunction.Match(Sheet2.Range("C5"), Sheet1.Range("W1:W438"), 0))
            Sheet1.Range("AA" & Sheet3.Range("B2")) = Sheet2.Range("E5")
            Sheet3.Range("B3").Value = CStr(WorksheetFunction.Match(Sheet2.Range("C5"), Sheet1.Range("W1:W438"), 0))
            Sheet1.Range("AC" & Sheet3.Range("B3")) = Now
        End If
    Next Email2
End Sub

Sub Contact_Update()
    Call Check_Update_Email
    Call Check_Update_Email2
End Sub
Sub New_Reset()
Attribute New_Reset.VB_ProcData.VB_Invoke_Func = " \n14"
'
' New_Reset Macro
'

'
    Range("C5").Select
    ActiveCell.FormulaR1C1 = ""
    Range("E5").Select
    ActiveCell.FormulaR1C1 = ""
    Range("C5").Select
End Sub
