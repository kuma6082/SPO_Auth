Private Sub Application_NewMailEx(ByVal EntryIDCollection As String) '受信時用
    Dim objItem As Object
    Set objItem = Session.GetItemFromID(EntryIDCollection)
    If TypeName(objItem) = "MailItem" Then
        SPO_Auth objItem
    End If
End Sub

Private Sub SPO_Auth_Mail(ByVal objItem As MailItem)
'CodeFolderはSPO_Authと同じフォルダ
Const CodeFolder = "*************"
mailad1 = "no-reply@notify.microsoft.com"
mailad2 = "no-reply@sharepointonline.com"

Dim filePath As String
Dim fileNo As Integer

Dim filePath2 As String
Dim fileNo2 As Integer

If objItem.Sender.Address Like mailad1 Or objItem.Sender.Address Like mailad2 Then
    objItem.UnRead = False
    filePath = CodeFolder & objItem.Subject & ".txt"
    fileNo = FreeFile
    Open filePath For Append As #fileNo
    Print #fileNo, ""
    Close #fileNo
    
    filePath2 = CodeFolder & Mid(objItem.body, InStr(objItem.body, """") + 1, 7) & "識別子.txt"
    fileNo = FreeFile
    Open filePath2 For Append As #fileNo
    Print #fileNo, ""
    Close #fileNo
End If
End Sub


