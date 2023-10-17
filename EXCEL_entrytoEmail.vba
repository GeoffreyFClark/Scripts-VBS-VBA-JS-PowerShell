' Example script: Create button to autosend email using info from a selected row. Row is selected via combobox dropdown list.

' (1) Open Excel worksheet.
' (2) "Developer" tab --> click "Insert"
' (3) Under "ActiveX Controls" --> click "ComboBox" then "Button" and place them on sheet
' (4) Right-click the ComboBox --> "Properties" --> Set "ListFillRange" to column letter containing representative selectable cells for each row (like "A:A" if you want to see cell entries of column A as dropdown options)
' (5) Right-click the Button --> "Properties" --> Set "Name" to "SendEmailButton".
' (6) Close Properties window, right-click Button --> click "View Code". VBA script:

Private Sub SendEmailButton_Click()

    Dim OutApp As Object
    Dim OutMail As Object
    Dim TargetRow As Long
    Dim EmailAddress As String
    Dim BodyText As String

    ' Get selected row number from ComboBox
    TargetRow = ComboBox1.Value ' (Rename "ComboBox1" here if you named it differently)

    ' Example of pulling email from second column's cell
    ' Rename "Sheet1" if different sheet name
    EmailAddress = ThisWorkbook.Sheets("Sheet1").Cells(TargetRow, 2).Value

    ' example using third column's cell value
    BodyText = "Your supply order has arrived. It includes: " & ThisWorkbook.Sheets("Sheet1").Cells(TargetRow, 3).Value

    ' Instantiate Outlook app object and new outgoing email
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    ' Set email properties
    With OutMail
        .To = EmailAddress
        .CC = "" ' Adjust as needed
        .BCC = "" ' Adjust as needed
        .Subject = "Supply Order Arrival" ' Adjust as needed
        .Body = BodyText
        ' .Attachments.Add "path" ' Can use this to add file attachment
        .Send 
        ' Can use .Display if you just want to display the email, not send it
        ' e.g. .Display the draft to allow custom modifications, then click send button (comment out .Send)
    End With

    ' Cleanup
    Set OutMail = Nothing
    Set OutApp = Nothing

End Sub