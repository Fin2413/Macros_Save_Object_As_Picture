Sub Save_Object_As_Picture()
    ' Declaring variables
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String, productName As String
    Dim topLeftCell As Range
    Dim logFilePath As String
    
    ' Setting the path to save images
    sImagesPath = ActiveWorkbook.Path & "\images\" ' The folder for saving images in the current directory of the book
    
    ' Setting the log file path in the images folder
    logFilePath = sImagesPath & "log.txt"
    
    ' Creating a folder if it does not exist
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath ' Create a folder for images if there is none
        LogMessage "Created folder: " & sImagesPath, logFilePath
    End If
    
    ' Disabling screen updates and warnings to speed up execution
    On Error Resume Next ' Ignore errors
    Application.ScreenUpdating = False ' Disable screen refresh
    Application.DisplayAlerts = False ' Disable warnings

    ' Installing the current sheet and creating a temporary sheet
    Set wsSh = ActiveSheet ' Installing the active sheet
    Set wsTmpSh = ActiveWorkbook.Sheets.Add ' Adding a time sheet to work with the schedule

    ' Iterating through all the objects on the active sheet
    For Each oObj In wsSh.Shapes
        ' Checking whether an object is an image
        If oObj.Type = 13 Then ' Type 13 are images
            li = li + 1 ' Counter for image names
            
            ' We get the cell where the upper left corner of the object is located
            Set topLeftCell = oObj.TopLeftCell
            
            ' We get the product name from the first column (column A) of the same row as the image
            productName = wsSh.Cells(topLeftCell.Row, 1).Value ' Name from column A
            
            ' Removing invalid characters from the file name
            productName = CleanFileName(productName)
            
            ' If the product name is empty, use the standard name
            If productName = "" Then
                productName = "img" & li
            End If
            
            ' Copying the image
            oObj.Copy

            ' Using a time graph to export an image
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0 ' Removing the boundaries of the graph
                .Parent.Select ' Choosing a schedule
                .Paste ' Inserting the image into the graph
                .Export Filename:=sImagesPath & productName & ".jpg", FilterName:="JPG" ' Exporting the image as a JPG file
                LogMessage "Saved image: " & productName & ".jpg", logFilePath ' Log saving image
                .Parent.Delete ' Deleting the time schedule after saving the image
            End With
            
            ' We write the file name to the cell where the image was located
            oObj.TopLeftCell.Value = productName ' We write the file name in the cell
        End If
    Next oObj

    ' Freeing up memory
    Set oObj = Nothing
    Set wsSh = Nothing
    wsTmpSh.Delete ' Deleting a temporary sheet

    ' Turning back the screen update and warnings
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ' Process completion message
    MsgBox "The objects are saved in the folder: " & sImagesPath, vbInformation, "Success"
    LogMessage "Process completed successfully.", logFilePath ' Log completion
End Sub

Function CleanFileName(fileName As String) As String
    ' Remove invalid characters from file name
    fileName = Replace(fileName, "/", "_")
    fileName = Replace(fileName, "\", "_")
    fileName = Replace(fileName, ":", "_")
    fileName = Replace(fileName, "*", "_")
    fileName = Replace(fileName, "?", "_")
    fileName = Replace(fileName, """", "_")
    fileName = Replace(fileName, "<", "_")
    fileName = Replace(fileName, ">", "_")
    fileName = Replace(fileName, "|", "_")
    
    ' Remove line breaks
    fileName = Replace(fileName, vbCr, "") ' Removing carriage return
    fileName = Replace(fileName, vbLf, "") ' Removing line feed
    fileName = Replace(fileName, vbCrLf, "") ' Removing carriage return + line feed
    
    CleanFileName = fileName
End Function

Sub LogMessage(message As String, logFilePath As String)
    Dim logFile As Integer
    logFile = FreeFile
    Open logFilePath For Append As #logFile
    Print #logFile, Now & ": " & message
    Close #logFile
End Sub
