Attribute VB_Name = "Module4"
Sub SNP_Check()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    ' Get the value from cell
    Dim Fw As String
    Dim Re As String
    Fw = ThisWorkbook.Sheets(1).Range("F2").Value
    Re = ThisWorkbook.Sheets(1).Range("G2").Value
    Ch = ThisWorkbook.Sheets(1).Range("B2").Value
    
    ' Open the browser and navigate to the desired URL
    ie.Visible = True
    ie.navigate "https://genetools.org/SNPCheck/snpcheck.htm"
    
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Ensure elements are loaded before interacting
    Do While ie.document.getElementById("primerPairText") Is Nothing
        DoEvents
    Loop
    
    ' Fill in the text fields and trigger change events
    With ie.document
        .getElementById("primerPairText").Value = "Fw_and_Re" & " " & Fw & " " & Re & " " & Ch
        .getElementById("primerPairText").FireEvent "onchange"
    End With
  
    ' Click the button with id="snpcheckButton_label"
    With ie.document
        .getElementById("snpcheckButton_label").Click
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
