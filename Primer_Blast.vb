Sub Primer_Blast()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
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
    
    sequence = ThisWorkbook.Sheets(1).Range("A2").Value
    min_prod = ThisWorkbook.Sheets("Variables").Range("B2").Value
    max_prod = ThisWorkbook.Sheets("Variables").Range("C2").Value
    min_tm = ThisWorkbook.Sheets("Variables").Range("D2").Value
    max_tm = ThisWorkbook.Sheets("Variables").Range("E2").Value
    max_diff_tm = ThisWorkbook.Sheets("Variables").Range("F2").Value
    database = ThisWorkbook.Sheets("Variables").Range("G2").Value
    max_gc = ThisWorkbook.Sheets("Variables").Range("H2").Value
    self = ThisWorkbook.Sheets("Variables").Range("I2").Value
    max_comp = ThisWorkbook.Sheets("Variables").Range("J2").Value
    
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
        
        .getElementById("PRIMER_PRODUCT_MIN").Value = min_prod
        .getElementById("PRIMER_PRODUCT_MIN").FireEvent "onchange"
        
        .getElementById("PRIMER_PRODUCT_MAX").Value = max_prod
        .getElementById("PRIMER_PRODUCT_MAX").FireEvent "onchange"
        
        .getElementById("PRIMER_MIN_TM").Value = min_tm
        .getElementById("PRIMER_MIN_TM").FireEvent "onchange"
        
        .getElementById("PRIMER_MAX_TM").Value = max_tm
        .getElementById("PRIMER_MAX_TM").FireEvent "onchange"
        
        .getElementById("PRIMER_MAX_DIFF_TM").Value = max_diff_tm
        .getElementById("PRIMER_MAX_DIFF_TM").FireEvent "onchange"
        
        .getElementById("PRIMER_SPECIFICITY_DATABASE").Value = database
        .getElementById("PRIMER_SPECIFICITY_DATABASE").FireEvent "onchange"
        
        .getElementById("PRIMER_MAX_GC").Value = max_gc
        .getElementById("PRIMER_MAX_GC").FireEvent "onchange"
        
        .getElementById("SELF_ANY").Value = self
        .getElementById("SELF_ANY").FireEvent "onchange"
        
        .getElementById("PRIMER_PAIR_MAX_COMPL_ANY").Value = max_comp
        .getElementById("PRIMER_PAIR_MAX_COMPL_ANY").FireEvent "onchange"
        
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
    
    ' Check if the final tab is loaded by looking for seq_1
    Dim finalTabLoaded As Boolean
    finalTabLoaded = False
    If Not ie.document.getElementById("seq_1") Is Nothing Then
        finalTabLoaded = True
    End If
    
    ' If final tab is not loaded, wait for it
    If Not finalTabLoaded Then
        ' Wait for the final tab to load
        Do While ie.Busy Or ie.readyState <> 4
            DoEvents
        Loop
    End If
    
    ' Check if the new elements exist before interacting
    If Not ie.document.getElementById("seq_1") Is Nothing Then
        ' Ensure the new elements are loaded before interacting
        Do While ie.document.getElementById("seq_1") Is Nothing Or ie.document.getElementById("nw1") Is Nothing
            DoEvents
        Loop
        
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
