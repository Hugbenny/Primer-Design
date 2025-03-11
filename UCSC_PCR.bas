Attribute VB_Name = "Module3"
Sub UCSC_PCR()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim Fw As String
    Dim Re As String
    Fw = ThisWorkbook.Sheets(1).Range("F2").Value
    Re = ThisWorkbook.Sheets(1).Range("G2").Value
    
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
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
