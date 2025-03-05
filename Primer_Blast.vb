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
    
    sequence = ActiveSheet.Range("A2").Value
    
    ' Set the active sheet and the Variables sheet
    Set wsActive = ActiveSheet
    Set wsVariables = ThisWorkbook.Sheets("Variables")
    
    ' Get the value in cell H2 of the active sheet
    searchValue = wsActive.Range("H2").Value
    
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
    Else
        ' Handle the case where the value is not found
        min_prod = "None"
        max_prod = "None"
        min_tm = "None"
        max_tm = "None"
        max_diff_tm = "None"
        database = "None"
        max_gc = "None"
        self = "None"
        max_comp = "None"
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
            .getElementById("PRIMER_DATABASE").Value = database
            .getElementById("PRIMER_DATABASE").FireEvent "onchange"
        End If
        If max_gc <> "None" Then
            .getElementById("PRIMER_MAX_GC").Value = max_gc
            .getElementById("PRIMER_MAX_GC").FireEvent "onchange"
        End If
        If self <> "None" Then
            .getElementById("PRIMER_SELF").Value = self
            .getElementById("PRIMER_SELF").FireEvent "onchange"
        End If
        If max_comp <> "None" Then
            .getElementById("PRIMER_MAX_COMP").Value = max_comp
            .getElementById("PRIMER_MAX_COMP").FireEvent "onchange"
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

' Loop to check each new tab loaded
Dim tabLoaded As Boolean
tabLoaded = True

Do While tabLoaded
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Check if the elements exist
    On Error Resume Next ' Ignore errors if elements do not exist
    If Not ie.document.getElementById("seq_1") Is Nothing And Not ie.document.getElementById("nw1") Is Nothing Then
        ' Check the boxes with id "seq_1" and "nw1"
        With ie.document
            .getElementById("seq_1").Checked = True
            .getElementById("seq_1").FireEvent "onclick"
            
            .getElementById("nw1").Checked = False
            .getElementById("nw1").FireEvent "onclick"
        End With
        
        ' Press the button with value "Submit"
        With ie.document
            Dim submitButton As Object
            Set submitButton = .querySelector("input[value='Submit']")
            If Not submitButton Is Nothing Then
                submitButton.Click
            End If
        End With
    End If
    On Error GoTo 0 ' Resume normal error handling
    
    ' Wait for another tab to be loaded
    ' (This part depends on how you determine if there are more tabs to load)
    ' For example, you might have a list of URLs to navigate to:
    ' If there are more URLs, navigate to the next one and set tabLoaded to True
    ' If not, set tabLoaded to False
    ' Example:
    ' If moreUrlsExist Then
    '     ie.Navigate nextUrl
    '     Do While ie.Busy Or ie.readyState <> 4
    '         DoEvents
    '     Loop
    '     tabLoaded = True
    ' Else
    '     tabLoaded = False
    ' End If
    
    ' For this example, we'll just exit the loop
    tabLoaded = False
Loop
    
    ' Uncomment to close the browser if needed
    ' ie.Quit
    ' Set ie = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
