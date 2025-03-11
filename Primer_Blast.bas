Attribute VB_Name = "Module1"
Sub Primer_Blast()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim wsActive As Worksheet
    Dim wsVariables As Worksheet
    Dim searchValue As String
    Dim foundCell As Range
    
    Dim sequence As String
    Dim min_prod As String
    Dim max_prod As String
    Dim min_tm As String
    Dim max_tm As String
    Dim max_diff_tm As String
    Dim database As String
    Dim max_gc As String
    Dim self As String
    Dim max_comp As String
    Dim max_target As String
        
    sequence = ActiveSheet.Range("A2").Value
    
    ' Set the active sheet and the Variables sheet
    Set wsActive = ActiveSheet
    Set wsVariables = ThisWorkbook.Sheets("Variables")
    
    ' Get the value in cell F2 of the active sheet
    searchValue = wsActive.Range("F2").Value
    
    ' Find the cell in the Variables sheet that matches the search value
    Set foundCell = wsVariables.Range("A:A").Find(What:=searchValue, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If the value is found, get the corresponding value in column
    If Not foundCell Is Nothing Then
        If Not IsEmpty(foundCell.Offset(0, 1).Value) Then min_prod = foundCell.Offset(0, 1).Value Else min_prod = "None"
        If Not IsEmpty(foundCell.Offset(0, 2).Value) Then max_prod = foundCell.Offset(0, 2).Value Else max_prod = "None"
        If Not IsEmpty(foundCell.Offset(0, 3).Value) Then min_tm = foundCell.Offset(0, 3).Value Else min_tm = "None"
        If Not IsEmpty(foundCell.Offset(0, 4).Value) Then max_tm = foundCell.Offset(0, 4).Value Else max_tm = "None"
        If Not IsEmpty(foundCell.Offset(0, 5).Value) Then max_diff_tm = foundCell.Offset(0, 5).Value Else max_diff_tm = "None"
        If Not IsEmpty(foundCell.Offset(0, 6).Value) Then database = foundCell.Offset(0, 6).Value Else database = "None"
        If Not IsEmpty(foundCell.Offset(0, 7).Value) Then max_gc = foundCell.Offset(0, 7).Value Else max_gc = "None"
        If Not IsEmpty(foundCell.Offset(0, 8).Value) Then self = foundCell.Offset(0, 8).Value Else self = "None"
        If Not IsEmpty(foundCell.Offset(0, 9).Value) Then max_comp = foundCell.Offset(0, 9).Value Else max_comp = "None"
        If Not IsEmpty(foundCell.Offset(0, 10).Value) Then max_target = foundCell.Offset(0, 10).Value Else max_target = "None"
    End If
    
    ' Open the browser and navigate to the desired URL
    ie.Visible = True
    ie.navigate "https://www.ncbi.nlm.nih.gov/tools/primer-blast/"
    
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
        
    ' Ensure elements are loaded before interacting
    Do While ie.document.getElementById("seq") Is Nothing
        DoEvents
    Loop
    
    ' Fill in the text fields and trigger change events
    With ie.document
        .getElementById("seq").Value = sequence
        .getElementById("seq").FireEvent "onchange"
        If min_prod <> "None" Then
            .getElementById("PRIMER_PRODUCT_MIN").Value = min_prod
            .getElementById("PRIMER_PRODUCT_MIN").FireEvent "onchange"
        End If
        If max_prod <> "None" Then
            .getElementById("PRIMER_PRODUCT_MAX").Value = max_prod
            .getElementById("PRIMER_PRODUCT_MAX").FireEvent "onchange"
        End If
        If min_tm <> "None" Then
            .getElementById("PRIMER_MIN_TM").Value = min_tm
            .getElementById("PRIMER_MIN_TM").FireEvent "onchange"
        End If
        If max_tm <> "None" Then
            .getElementById("PRIMER_MAX_TM").Value = max_tm
            .getElementById("PRIMER_MAX_TM").FireEvent "onchange"
        End If
        If max_diff_tm <> "None" Then
            .getElementById("PRIMER_MAX_DIFF_TM").Value = max_diff_tm
            .getElementById("PRIMER_MAX_DIFF_TM").FireEvent "onchange"
        End If
        If database <> "None" Then
            .getElementById("PRIMER_SPECIFICITY_DATABASE").Value = database
            .getElementById("PRIMER_SPECIFICITY_DATABASE").FireEvent "onchange"
        End If
        If max_gc <> "None" Then
            .getElementById("PRIMER_MAX_GC").Value = max_gc
            .getElementById("PRIMER_MAX_GC").FireEvent "onchange"
        End If
        If self <> "None" Then
            .getElementById("SELF_ANY").Value = self
            .getElementById("SELF_ANY").FireEvent "onchange"
        End If
        If max_comp <> "None" Then
            .getElementById("PRIMER_PAIR_MAX_COMPL_ANY").Value = max_comp
            .getElementById("PRIMER_PAIR_MAX_COMPL_ANY").FireEvent "onchange"
        End If
        If max_target <> "None" Then
            .getElementById("NUM_TARGETS").Value = max_target
            .getElementById("NUM_TARGETS").FireEvent "onchange"
        End If
        .getElementById("NO_SNP").Checked = True
        .getElementById("NO_SNP").FireEvent "onclick"
        
        .getElementById("nw1").Checked = False
        .getElementById("nw1").FireEvent "onclick"
        
        .getElementById("show_sviewer1").Checked = False
        .getElementById("show_sviewer1").FireEvent "onclick"
        
        .getElementById("nw2").Checked = False
        .getElementById("nw2").FireEvent "onclick"
        
        .getElementById("show_sviewer2").Checked = False
        .getElementById("show_sviewer2").FireEvent "onclick"
    End With
    
    ' Ensure the buttons are loaded before clicking
    Do While ie.document.getElementsByClassName("blastbutton prbutton").Length < 2
        DoEvents
    Loop
    
    ' Click the second button with class "blastbutton prbutton"
    With ie.document.getElementsByClassName("blastbutton prbutton")(1)
        .Focus
        .Click
    End With
    
    ' Wait for the new tab to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop

    Dim pairs_number As Integer
    Dim pairs_number_str As String
    
    ' Ensure the element is loaded before interacting
    Do While ie.document.getElementsByName("PRIMER_PAIRS_NUMBER").Length = 0
        DoEvents
    Loop
    
    ' Retrieve the value of the element as a string
    pairs_number_str = ie.document.getElementsByName("PRIMER_PAIRS_NUMBER")(0).Value
    
    ' Convert the string to an integer
    pairs_number = CInt(pairs_number_str)

    ' Output the value in a message box for debugging purposes
    ActiveSheet.Cells(2, 7).Value = pairs_number ' Write pairs_number value to column G2
'    MsgBox "Pairs Number: " & pairs_number

    Dim i As Integer
    Dim j As Integer
'    Dim table As Object
'    Dim row As Object
'    Dim cell As Object

'    ' Find the table element
'    Set table = ie.document.getElementsByTagName("table")(0) ' Adjust the index if there are multiple tables
'
'    ' Loop through the table rows and cells
'    i = 20 'Row
'    For Each row In table.Rows
'        j = 1 'Column
'        For Each cell In row.Cells
'            ActiveSheet.Cells(i, j).Value = cell.innerText
'            j = j + 1
'        Next cell
'        i = i + 1
'    Next row

    Dim fw_primer As String
    Dim fw_tm As String
    Dim rv_primer As String
    Dim rv_tm As String
    Dim product_length As String
    
    ' Loop through each pair and retrieve the values
    For i = 0 To pairs_number - 1
        ' Ensure the elements are loaded before interacting
        Do While ie.document.getElementsByName("FW_PRIMER_SEQ_" & i).Length = 0 Or ie.document.getElementsByName("RV_PRIMER_SEQ_" & i).Length = 0
            DoEvents
        Loop

        ' Retrieve the values
        fw_primer = ie.document.getElementsByName("FW_PRIMER_SEQ_" & i)(0).Value
        fw_tm = ie.document.getElementsByName("FW_PRIMER_TM_" & i)(0).Value
        rv_primer = ie.document.getElementsByName("RV_PRIMER_SEQ_" & i)(0).Value
        rv_tm = ie.document.getElementsByName("RV_PRIMER_TM_" & i)(0).Value
        product_length = ie.document.getElementsByName("PRODUCT_LENGTH_" & i)(0).Value
        
        ' Write the values to the active sheet, starting from row 2
        ActiveSheet.Cells(i + 9, 1).Value = fw_primer ' Write FW_PRIMER_SEQ_ value to column A9
        ActiveSheet.Cells(i + 9, 2).Value = fw_tm ' Write FW_PRIMER_TM_ value to column B9
        ActiveSheet.Cells(i + 9, 3).Value = rv_primer ' Write RV_PRIMER_SEQ_ value to column C9
        ActiveSheet.Cells(i + 9, 4).Value = rv_tm ' Write RV_PRIMER_TM_ value to column D9
        ActiveSheet.Cells(i + 9, 5).Value = product_length ' Write PRODUCT_LENGTH_ value to column D9
    Next i

    ' Exit the subroutine if no errors occur
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
