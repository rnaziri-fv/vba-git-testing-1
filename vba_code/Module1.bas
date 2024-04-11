Attribute VB_Name = "Module1"
Function SumNumbers(Number1 As Double, Number2 As Double) As Double
    ' Line change here.
    ' This function returns the sum of two numbers.
    SumNumbers = Number1 + Number2
End Function

Sub InsertCurrentDateTime()
    ' This macro inserts the current date and time into cell A1 of the active sheet.
    With ActiveSheet.Range("A1")
        .Value = Now
        .NumberFormat = "mm/dd/yyyy hh:mm:ss AM/PM"
    End With
End Sub
