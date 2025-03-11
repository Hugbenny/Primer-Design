Attribute VB_Name = "Module5"
Sub New_Primers()
    ' Declare a variable to hold the active sheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    ' Create a copy of the active sheet
    ws.Copy After:=ws
    
    ' Optionally, you can rename the copied sheet
    ' ActiveSheet.Name = ws.Name & " Copy"
End Sub
Sub Reset()
    ' Clear the contents of the specified cells on the active sheet
    With ActiveSheet
        .Range("A2").ClearContents
        .Range("B2").ClearContents
        .Range("C2").ClearContents
        .Range("D2").ClearContents
        .Range("E2").ClearContents
        .Range("F2").ClearContents
        .Range("G2").ClearContents
    End With
End Sub
