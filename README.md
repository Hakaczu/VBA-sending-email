# Visual Basic for Application - Email sendgin macro for Excel
This VBA code are simple example for email sending macro.

## Function for send email
`Public Function Send(body As String, adress As String, mailTitle As String) 'Function send email

    Dim outlookApp As Outlook.Application
    Dim myMail As Outlook.MailItem
    
    Set outlookApp = New Outlook.Application 'Tworzenie instancji obiektu Outlook.Application
    Set myMail = outlookApp.CreateItem(olMailItem) 'Tworzenie instancji miala
    
    myMail.To = adress
    myMail.Subject = mailTitle
    'html mail styling
    myMail.HTMLBody = "<html> <head> <style> table, th, td {border: 1px solid black;  border-collapse: collapse; padding: 3px; text-align: center;} table {width: 100%;}, important {background-color: yellow;}</style> </head> <body>" + body + "</body> </html>"
    myMail.Send 'Sending Email

End Function`

Works on Excel 2016
License Apache 2.0, please keep info about author. Open for commercial use and editing. 
