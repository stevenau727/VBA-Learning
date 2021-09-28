Attribute VB_Name = "Module1"
Sub Fastkey()
Attribute Fastkey.VB_ProcData.VB_Invoke_Func = "Z\n14"

Application.OnKey "+^{S}", "a"
Application.OnKey "+^{B}", "CreateMail"
End Sub

Sub a()

MsgBox ("hihi")


End Sub

Sub CreateMail()

    Dim olApp As Outlook.Application
    Set olApp = New Outlook.Application
    
    
    Dim olMail As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    
    olMail.Display
    
    MsgBox "Outlook opened"
    
    Set olMail = Nothing
    Set olApp = Nothing
    
End Sub

Sub SendMail()

    Dim olApp As Outlook.Application
    Set olApp = New Outlook.Application
    
    Dim olMail As Outlook.MailItem
    Set olMail = olApp.CreateItem(olMailItem)
    
    With olMail
        .To = "waitoau@gmail.com"
        
        .Subject = "Hello"
        
        .Body = "This is email content."
        
        .BodyFormat = olFormatPlain
    End With
    
    olMail.Attachments.Add ThisWorkbook.Path & "\trail.docx"
    
    'olMail.Save
    olMail.Display
    'olMail.Send
    
    Set olMail = Nothing
    Set olApp = Nothing
    
End Sub




