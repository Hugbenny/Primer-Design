Attribute VB_Name = "Module3"
Sub UCSC_PCR()
    On Error GoTo ErrorHandler
    
    ' Get the value from cell
    Dim Fw As String
    Dim Re As String
    Dim link As Object
    Dim href As String
    Dim linkText As String
    Dim row As Integer
    
    ' Loop through the rows in the Excel sheet
    row = 9 ' Start from the nineth row
    Do While Not IsEmpty(ActiveSheet.Range("A" & row).Value) And Not IsEmpty(ActiveSheet.Range("C" & row).Value)
        ' Get the values from the cells
        Fw = ActiveSheet.Range("A" & row).Value
        Re = ActiveSheet.Range("C" & row).Value
        
        ' Create Internet Explorer object
        Dim ie As Object
        Set ie = CreateObject("InternetExplorer.Application")
        
        ' Open the browser and navigate to the desired URL
        ie.Visible = True
        ie.navigate "https://genome-euro.ucsc.edu/cgi-bin/hgPcr?hgsid=347282789_EpuHMdrAxrBnnrPR4FAGokTOrAMR"
        
        ' Wait for the page to fully load
        Do While ie.Busy Or ie.readyState <> 4
            DoEvents
        Loop
            
        ' Ensure elements are loaded before interacting
        Do While ie.document.getElementById("wp_f") Is Nothing Or ie.document.getElementById("wp_r") Is Nothing
            DoEvents
        Loop
        
        ' Fill in the text fields and trigger change events
        With ie.document
            .getElementById("wp_f").Value = Fw
            .getElementById("wp_f").FireEvent "onchange"
            .getElementById("wp_r").Value = Re
            .getElementById("wp_r").FireEvent "onchange"
        End With
        
        ' Click the button with id="Submit"
        With ie.document
            .getElementById("Submit").Click
        End With
        
        ' Wait for the page to fully load
        Do While ie.Busy Or ie.readyState <> 4
            DoEvents
        Loop
        
        ' Ensure the element is loaded before interacting
        Do While ie.document.getElementsByTagName("a").Length = 0
            DoEvents
        Loop
    
        ' Find the <a> element
        Set link = ie.document.querySelector("a[href*='chr']") ' Adjust the selector as needed
    
        ' Get the href attribute and inner text
        href = link.href
        linkText = link.innerText
    
        ' Write the information to the Excel sheet
        ActiveSheet.Cells(row, 8).Value = href
        ActiveSheet.Cells(row, 7).Value = linkText
    
        ' Clean up
        ie.Quit
        Set ie = Nothing
        Set link = Nothing

        ' Move to the next row
        row = row + 1
    Loop
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
