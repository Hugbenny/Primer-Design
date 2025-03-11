Attribute VB_Name = "Order"
Sub Order()
    On Error GoTo ErrorHandler
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")

    ' Construct the URL
    Dim url As String
    url = "https://rick.nerial.uk"
    
    ' Open the browser and navigate to the desired URL
    ie.Visible = True
    ie.navigate url
    
    ' Wait for the page to fully load
    Do While ie.Busy Or ie.readyState <> 4
        DoEvents
    Loop
    
    ' Additional code to interact with the webpage can be added here
    
    ' Exit the subroutine to avoid running the error handler
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
End Sub
