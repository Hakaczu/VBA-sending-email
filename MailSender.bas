'Author Sebastian Szypulski vel Sisa

Public Function Headers() As String 'Function return headers for table in email. 

    Dim column As String 'Name of column
    Headers = "<tr>"
    For k = 1 To 18 Step 1

        column = Sheets(1).Cells(1, k)
        Headers = Headers + "<th> " + column + " </th>" 'Getting table headers from Sheet 

    Next k
    Headers = Headers + "</tr>" 'return headers

End Function
Public Function Product(row As Integer) As String 'Function return product for table in email

    Dim column As String
    Product = "<tr>"
    For k = 1 To 18 Step 1

        column = Sheets(1).Cells(row, k)
        If k = 12 Then
            Product = Product + "<td style= 'background-color: yellow' > " + column + " </td>" 'column distinction
        Else
            Product = Product + "<td> " + column + " </td>"
        End If
    Next k
    Product = Product + "</tr>" 'return product

End Function
Public Function GetAdress(brand As String) 'Function return adress for email
    Dim rowHow As Integer
    Dim colHow As Integer
    Dim rowID As Integer

    rowID = 0
    rowHow = Sheets(2).UsedRange.Rows.count
    colHow = Sheets(2).UsedRange.Columns.count

    For i = 1 To rowHow Step 1
        If CStr(Sheets(2).Cells(i, 1).Value) = CStr(brand) Then
            rowID = i
        End If
    Next i

    GetAdress = ""
    If rowID = 0 Then
        MsgBox "I can't find email for this brand: " + brand
    Else
        For i = 1 To colHow Step 1
            If IsEmpty(Sheets(2).Cells(rowID, i + 1)) = False Then
                GetAdress = GetAdress + CStr(Sheets(2).Cells(rowID, i + 1).Value) + ";"
            End If
        Next i
    End If
End Function
Public Function mailTitle(mailType As String, row As Integer) 'Function return mail subject

    If mailType = "block" Then
        mailTitle = CStr(Sheets(1).Cells(row, 1).Value) + "-" + CStr(Sheets(1).Cells(row, 7).Value) + "-BLOCK WMS"
    End If
    If mailType = "unlock" Then
        mailTitle = CStr(Sheets(1).Cells(row, 1).Value) + "-" + CStr(Sheets(1).Cells(row, 7).Value) + "-UNLOCK WMS"
    End If

End Function
Public Function Send(body As String, adress As String, mailTitle As String) 'Function send email

    Dim outlookApp As Outlook.Application
    Dim myMail As Outlook.MailItem
    
    Set outlookApp = New Outlook.Application
    Set myMail = outlookApp.CreateItem(olMailItem)
    
    myMail.To = adress
    myMail.Subject = mailTitle
    'html mail styling
    myMail.HTMLBody = "<html> <head> <style> table, th, td {border: 1px solid black;  border-collapse: collapse; padding: 3px; text-align: center;} table {width: 100%;}, important {background-color: yellow;}</style> </head> <body>" + body + "</body> </html>"
    myMail.Send 'Sending Email

End Function
Public Function BlockBody(row As Integer, range As Integer) As String 'Mail body

    Dim i As Integer
    Dim endRange As Integer
    endRange = row + range - 1
    BlockBody = "<p>Model <b>" + CStr(Sheets(1).Cells(row, 1).Value) + "</b> han been block</p><br>"
    BlockBody = BlockBody + "<table>" + Headers()
    For i = row To endRange Step 1
        BlockBody = BlockBody + Product(i)
    Next i
    BlockBody = BlockBody + "</table>"
    BlockBody = BlockBody + "Regards Quality Control"

End Function
Public Function UnlockBody(row As Integer, range As Integer) As String 'Mail body

    Dim i As Integer
    Dim endRange As Integer
    endRange = row + range - 1
    UnlockBody = "<p>Model <b>" + CStr(Sheets(1).Cells(row, 1).Value) + "</b> has been unlock.</p><br>"
    UnlockBody = UnlockBody + "<table>" + Headers()
    For i = row To endRange Step 1
        UnlockBody = UnlockBody + Product(i)
    Next i
    UnlockBody = UnlockBody + "</table>"
    UnlockBody = UnlockBody + "Regards Quality Control"

End Function
Public Function CanSend(row As Integer, range As Integer, mailType As String) As Boolean 'check that can send email
    Dim endRange As Integer
    endRange = row + range - 1
    CanSend = True
    If mailType = "block" Then
        For i = row To endRange Step 1
            If IsEmpty(Sheets(1).Cells(i, 30)) = False Then
                CanSend = False
            End If
        Next i
    End If
    If mailType = "unlock" Then
        For i = row To endRange Step 1
            If CStr(Sheets(1).Cells(i, 30)) = "odblokowano" Or IsEmpty(Sheets(1).Cells(i, 30)) = True Then
                CanSend = False
            End If
        Next i
    End If
End Function
Public Function WriteInfo(row As Integer, range As Integer, infoType As String) 'write info about email status
    Dim endRange As Integer
    endRange = row + range - 1
    
    If infoType = "block" Then
        For i = row To endRange Step 1
            Sheets(1).Cells(i, 30) = "block"
        Next i
    End If
    If infoType = "unlock" Then
        For i = row To endRange Step 1
            Sheets(1).Cells(i, 30) = "unlock"
        Next i
    End If
    
End Function
Sub BlockMail()
    If ActiveWorkbook.ReadOnly Then 'protection against unwanted use
        Exit Sub
    End If
    Dim body As String
    Dim adress As String
    Dim row As Integer
    Dim brand As String
    Dim mailSubject As String
    Dim range As Integer
    Dim mailType As String
    
    mailType = "block"
    range = Selection.Rows.count
    row = ActiveCell.row
    
    If CanSend(row, range, mailType) = True Then
        brand = Sheets(1).Cells(row, 7).Value
        adress = GetAdress(brand)
        body = BlockBody(row, range)
        mailSubject = mailTitle(mailType, row)
        
        Call Send(body, adress, mailSubject)
        Call WriteInfo(row, range, mailType)
    Else
        MsgBox "Block Mail has been sent"
    End If
End Sub
Sub UnlockMail()
If ActiveWorkbook.ReadOnly Then 'protection against unwanted use
        Exit Sub
    End If
    Dim body As String
    Dim adress As String
    Dim row As Integer
    Dim brand As String
    Dim mailSubject As String
    Dim range As Integer
    Dim mailType As String
    
    mailType = "unlock"
    range = Selection.Rows.count
    row = ActiveCell.row
    
    If CanSend(row, range, mailType) = True Then
        brand = Sheets(1).Cells(row, 7).Value
        adress = GetAdress(brand)
        body = UnlockBody(row, range)
        mailSubject = mailTitle(mailType, row)
        
        Call Send(body, adress, mailSubject)
        Call WriteInfo(row, range, mailType)
    Else
        MsgBox "Unlock Mail has been sent"
    End If
End Sub
