Attribute VB_Name = "Others"
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
    
    Dim tableRange As Range
    Dim cell As Range
    Dim shp As Shape
    
    ' Clear the contents of the specified cells on the active sheet
    With ActiveSheet
        .Range("A2:E2").ClearContents
        .Range("H2:I2").ClearContents
    End With
    
    ' Set the range of the table
    Set tableRange = ActiveSheet.Range("A9:M18") ' Replace with your table range

    ' Clear the contents of the table
    tableRange.ClearContents
    
    ' Loop through each shape in the worksheet
    For Each shp In ActiveSheet.Shapes
        ' Check if the shape is within the table range
        If Not Intersect(shp.TopLeftCell, tableRange) Is Nothing Then
            shp.Delete
        End If
    Next shp
    
End Sub
