Attribute VB_Name = "Module4"
Sub SNP_Check()
    On Error GoTo ErrorHandler
   
    ' Get the value from cell
    Dim element As Object
    Dim img As Object
    Dim imgURL As String
    Dim img_primer As Object
    Dim imgURL_primer As String
    Dim cell As Range
    Dim row As Integer
    Dim border As Variant
    Dim Fw As String
    Dim Re As String
    Dim Ch As String
    
    ' Create Internet Explorer object
    Dim ie As Object
    Dim newIE  As Object
    
    Ch = ActiveSheet.Range("B2").Value
        
    ' Loop through the rows in the Excel sheet
    row = 9 ' Start from the nineth row
    Do While Not IsEmpty(ActiveSheet.Range("A" & row).Value) And Not IsEmpty(ActiveSheet.Range("C" & row).Value)
        ' Get the values from the cells
        Fw = ActiveSheet.Range("A" & row).Value
        Re = ActiveSheet.Range("C" & row).Value
                
        ' Open the browser and navigate to the desired URL
        Set ie = CreateObject("InternetExplorer.Application")
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
        
        ' Wait for the page to fully load the results
        Do While ie.Busy Or ie.readyState <> 4
            DoEvents
        Loop
               
'        ' Add a 5-second timer
'        Application.Wait (Now + TimeValue("0:00:05"))
                   
        ' Wait for the results URL to be fully loaded and not the loading image URL
        Dim startTime As Double
        startTime = Timer
        Do
            Set img = ie.document.getElementById("Fw_and_Re.result").getElementsByTagName("img")(0)
            If Not img Is Nothing Then ' Check if img exists
                imgURL = img.src
            Else
                imgURL = ""
            End If
            DoEvents
            ' Add a small delay
            Application.Wait (Now + TimeValue("0:00:01"))
            ' Timeout after 30 seconds
            If Timer - startTime > 30 Then
                MsgBox "Timeout waiting for the image URL to update"
                Exit Do
            End If
        Loop While imgURL = "https://genetools.org/SNPCheck/img/working.gif"
        
        ' Check for error message
        If ie.document.getElementById("Fw_and_Re.result").innerText Like "*Error*" Then
'            MsgBox "Error detected: No hits found within 5000 bps"
            ' Write "Error" in the corresponding cell
            With ActiveSheet.Cells(row, 6)
                .Value = "Error"
                .Font.Bold = True
                .Font.Color = RGB(255, 0, 0)
            End With
            ' Move to the next row
            row = row + 1
            ' Close the main browser
            ie.Quit
            Set ie = Nothing
            ' Continue to the next iteration
            GoTo ContinueLoop
        End If
                   
        ' Find the image element within the specific div
        If Not ie.document.getElementById("Fw_and_Re.result") Is Nothing Then
            Set img = ie.document.getElementById("Fw_and_Re.result").getElementsByTagName("img")(0)

            ' Get the image URL
            imgURL = img.src
            Debug.Print "Image URL: " & imgURL

'            ' Wait for the results URL to be fully loaded and not the loading image URL
'            Do While imgURL = "https://genetools.org/SNPCheck/img/working.gif"
'                imgURL = img.src
'            Loop

            ' Open a new Internet Explorer instance
            Set newIE = CreateObject("InternetExplorer.Application")
            newIE.Visible = True
            newIE.navigate imgURL

            ' Wait for the new page to load
            Do While newIE.Busy Or newIE.readyState <> 4
                DoEvents
            Loop

            ' Copy the image to the clipboard using the new Internet Explorer instance
            newIE.ExecWB 17, 0 ' Select All
            newIE.ExecWB 12, 2 ' Copy

            Set cell = ActiveSheet.Cells(row, 6)

'            ' Store the current border properties
'            For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
'                cell.Borders(border).LineStyle = cell.Borders(border).LineStyle
'                cell.Borders(border).Weight = cell.Borders(border).Weight
'                cell.Borders(border).Color = cell.Borders(border).Color
'            Next border

            ' Insert the image into the Excel sheet using Shapes.AddPicture
            ActiveSheet.Shapes.AddPicture imgURL, _
                msoFalse, msoCTrue, _
                cell.Left, cell.Top, 15, 15   ' Adjust the width and height as needed

'            ' Reapply the stored border properties
'            For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
'                cell.Borders(border).LineStyle = cell.Borders(border).LineStyle
'                cell.Borders(border).Weight = cell.Borders(border).Weight
'                cell.Borders(border).Color = cell.Borders(border).Color
'            Next border

            ' Close the browser
            newIE.Quit
            Set newIE = Nothing
         End If

        ' Find the second image element within the specific div
        If ie.document.getElementsByClassName("primer_img").Length > 0 Then
            Set img_primer = ie.document.getElementById("Fw_and_Re.img").getElementsByTagName("img")(0)

            ' Get the image URL
            imgURL_primer = img_primer.src
            Debug.Print "Image URL: " & imgURL_primer

            ' Open a new Internet Explorer instance
            Set newIE = CreateObject("InternetExplorer.Application")
            newIE.Visible = True
            newIE.navigate imgURL_primer
            
            ' Wait for the new page to load
            Do While newIE.Busy Or newIE.readyState <> 4
                DoEvents
            Loop
            
            ' Copy the image to the clipboard using the new Internet Explorer instance
            newIE.ExecWB 17, 0 ' Select All
            newIE.ExecWB 12, 2 ' Copy
            
            Set cell = ActiveSheet.Cells(row, 10)

'            ' Store the current border properties
'            For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
'                cell.Borders(border).LineStyle = cell.Borders(border).LineStyle
'                cell.Borders(border).Weight = cell.Borders(border).Weight
'                cell.Borders(border).Color = cell.Borders(border).Color
'            Next border

            ' Insert the image into the Excel sheet using Shapes.AddPicture
            Dim originalWidth As Single
            Dim originalHeight As Single
            Dim newWidth As Single
            Dim newHeight As Single
            Dim maxHeight As Single

            ' Set the maximum height
            maxHeight = 15

            ' Get the original dimensions of the image
            originalWidth = newIE.document.images(0).Width
            originalHeight = newIE.document.images(0).Height

            ' Calculate the new dimensions while maintaining the aspect ratio based on height
            newHeight = maxHeight
            newWidth = (originalWidth / originalHeight) * maxHeight

            ' Insert the image with the new dimensions
            With ActiveSheet.Shapes.AddPicture(imgURL_primer, _
                msoFalse, msoCTrue, _
                cell.Left, cell.Top, newWidth, newHeight)
                .LockAspectRatio = msoTrue ' Lock the aspect ratio
            End With
            
'            ' Reapply the stored border properties
'            For Each border In Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal)
'                cell.Borders(border).LineStyle = cell.Borders(border).LineStyle
'                cell.Borders(border).Weight = cell.Borders(border).Weight
'                cell.Borders(border).Color = cell.Borders(border).Color
'            Next border
            
            ' Close the browser
            newIE.Quit
            Set newIE = Nothing
        End If
        
        ' Clean up
        Set img = Nothing
        imgURL = ""
        Set img_primer = Nothing
        imgURL_primer = ""
        
        ' Close the browser
        ie.Quit
        Set ie = Nothing
               
        ' Move to the next row
        row = row + 1
ContinueLoop:
    Loop
    
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
    If Not ie Is Nothing Then
        ie.Quit
        Set ie = Nothing
    End If
    If Not newIE Is Nothing Then
        newIE.Quit
        Set newIE = Nothing
    End If
    Resume ContinueLoop
End Sub
