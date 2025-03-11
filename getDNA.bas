Attribute VB_Name = "getDNA"
Sub getDNA_Coordinates()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim Ch As String
    Dim St As String
    Dim En As String
    Dim genome As String
    
    Ch = ActiveSheet.Range("B2").Value
    St = ActiveSheet.Range("C2").Value
    En = ActiveSheet.Range("D2").Value
    genome = ActiveSheet.Range("G2").Value
    
    ' Construct the URL with the position parameter
    Dim url As String
    url = "https://genome.ucsc.edu/cgi-bin/hgc?g=getDna&i=mixed&c=chr" & Ch & "&l=" & St & "&r=" & En & "&db=" & genome
    
    ' Open the browser and navigate to the desired URL
    ie.Visible = True
    ie.navigate url
    
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Ensure elements are loaded before interacting
    Do While ie.document.getElementById("getDnaPos") Is Nothing
        DoEvents
    Loop
    
    ' Click the button with id="goButton"
    With ie.document
        .getElementById("submit").Click
    End With
    
    ' Wait for the new tab to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Extract the content from the <pre> tag
    Dim content As String
    content = ie.document.getElementsByTagName("pre")(0).innerText ' Assuming the content is in the first <pre> tag
    
    ' Paste the content into Excel
    ActiveSheet.Range("A2").Value = content
    
    ' Create a link to UCSC
    ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(2, 9), Address:="https://genome-euro.ucsc.edu/cgi-bin/hgTracks?db=" & genome & "&lastVirtModeType=default&lastVirtModeExtraState=&virtModeType=default&virtMode=0&nonVirtPosition=&position=chr" & Ch & "%3A" & St & "-" & En, TextToDisplay:="UCSC"
    
    ' Uncomment to close the browser if needed
    ie.Quit
    Set ie = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
Sub getDNA_Gene()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim Ge As String
    Dim genome As String
    
    Ge = ActiveSheet.Range("E2").Value
    genome = ActiveSheet.Range("G2").Value
    
    ' Construct the URL with the position parameter
    Dim url As String
    url = "https://genome.ucsc.edu/cgi-bin/hgGene?db=" & genome & "&hgg_gene=" & Ge
    
    ' Open the browser and navigate to the desired URL
    ie.Visible = True
    ie.navigate url
    
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Ensure the links are loaded before interacting
    Do While ie.document.getElementsByTagName("a").Length = 0
        DoEvents
    Loop
    
    ' Loop through all links and click the one containing "Genomic Sequence"
    Dim link As Object
    Dim sequence As String
    Dim Ch As String
    Dim St As String
    Dim En As String
    For Each link In ie.document.getElementsByTagName("a")
        If InStr(link.innerText, "Genomic Sequence") > 0 Then
            ' Extract the sequence between parentheses
            sequence = Mid(link.innerText, InStr(link.innerText, "(") + 1, InStr(link.innerText, ")") - InStr(link.innerText, "(") - 1)
            
            ' Extract the chromosome number (Ch)
            Ch = Mid(sequence, InStr(sequence, "chr") + 3, InStr(sequence, ":") - InStr(sequence, "chr") - 3)
            
            ' Extract the start position (St)
            St = Mid(sequence, InStr(sequence, ":") + 1, InStr(sequence, "-") - InStr(sequence, ":") - 1)
            
            ' Extract the end position (En)
            En = Mid(sequence, InStr(sequence, "-") + 1, Len(sequence) - InStr(sequence, "-"))
            
            ' Click the link
            link.Click
            Exit For
        End If
    Next link
    
    ' Wait for the new content to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Click the button with id="goButton"
    With ie.document
        .getElementById("submit").Click
    End With
    
    ' Wait for the new tab to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Extract the content from the <pre> tag
    Dim content As String
    content = ie.document.getElementsByTagName("pre")(0).innerText ' Assuming the content is in the first <pre> tag
    
    ' Paste the content into Excel
    ActiveSheet.Range("A2").Value = content
    ActiveSheet.Range("B2").Value = Ch
    ActiveSheet.Range("C2").Value = St
    ActiveSheet.Range("D2").Value = En
        
    ' Create a link to UCSC
    ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(2, 9), Address:="https://genome-euro.ucsc.edu/cgi-bin/hgTracks?db=" & genome & "&lastVirtModeType=default&lastVirtModeExtraState=&virtModeType=default&virtMode=0&nonVirtPosition=&position=chr" & Ch & "%3A" & St & "-" & En, TextToDisplay:="UCSC"
        
    ' Uncomment to close the browser if needed
    ie.Quit
    Set ie = Nothing
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
