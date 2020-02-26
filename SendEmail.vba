Sub Send_Email()

Dim OutApp, OutMail As Object
Dim Filepath, Imgpath, ImgName As String
Dim TempFilePath As String

Set OutApp = CreateObject("Outlook.Application")
Set OutMail = OutApp.CreateItem(olMailItem)

Filepath = ActiveWorkbook.FullName
ImgName = "ImageName" & ".png"
Imgpath = ActiveWorkbook.Path & "\" & ImgName

With OutMail
    .To = ""
    .SentOnBehalfOfName = ""
    .Cc = ""
    .BCC = ""
    .Subject = "Email Subject" & Format(Now(), "mmddyyyy")
    
    TempFilePath = Environ$("temp") & "\IMG.png" 
    .Attachments.Add TempFilePath, olByValue, 0 'attach image file from workbook
    .Attachments.Add Filepath 'attach file from path 

    .HTMLBody = "<html>Hello,<br><br>" & _
            "Email Message." & _
            "<br><br>" & _
            "<img src='cid:IMG.png' width='600'>" & _
            "<br><br>Regards,<br>" & _
            "<B>Bold Signature</B><br><br>"
    .Send
    '.Display
End With

On Error GoTo 0

Set OutMail = Nothing
Set OutApp = Nothing

End Sub
