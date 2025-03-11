Attribute VB_Name = "Module2"
Sub getDNA_Coordinates()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim Ch As String
    Dim St As String
    Dim En As String
    Ch = ThisWorkbook.Sheets(1).Range("B2").Value
    St = ThisWorkbook.Sheets(1).Range("C2").Value
    En = ThisWorkbook.Sheets(1).Range("D2").Value
    
    ' Construct the URL with the position parameter
    Dim url As String
    url = "https://genome.ucsc.edu/cgi-bin/hgc?g=getDna&i=mixed&c=chr" & Ch & "&l=" & St & "&r=" & En & "&db=hg19"
    
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
    ThisWorkbook.Sheets(1).Range("A2").Value = content
    
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
    Ge = ThisWorkbook.Sheets(1).Range("E2").Value
    
    ' Construct the URL with the position parameter
    Dim url As String
    url = "https://genome.ucsc.edu/cgi-bin/hgGene?db=hg19&hgg_gene=" & Ge
    
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
    For Each link In ie.document.getElementsByTagName("a")
        If InStr(link.innerText, "Genomic Sequence") > 0 Then
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
    ThisWorkbook.Sheets(1).Range("A2").Value = content
    
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
