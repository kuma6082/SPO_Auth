VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'保存用
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
 '               (ByVal hwnd As Long, ByVal lpszOp As String, _
  '               ByVal lpszFile As String, ByVal lpszParams As String, _
   '              ByVal LpszDir As String, ByVal FsShowCmd As Long) _
    '             As Long
'
Private Sub Application_NewMailEx(ByVal EntryIDCollection As String) '受信時用
    Dim objItem As Object
    Set objItem = Session.GetItemFromID(EntryIDCollection)
    If TypeName(objItem) = "MailItem" Then
        payroll objItem
        SaveAttach objItem
        Date_Posting_DDPS objItem
        LinkCheck objItem
        SPO_Auth objItem
        reply_drawing objItem
    End If
End Sub
Private Sub SPO_Auth(ByVal objItem As MailItem)
Const SAVE_PATH = "D:\@VBA\007_SPO操作\PASS\"
mailad1 = "no-reply@notify.microsoft.com"
mailad2 = "no-reply@sharepointonline.com"

Dim filePath As String
Dim fileNo As Integer

Dim filePath2 As String
Dim fileNo2 As Integer

If objItem.Sender.Address Like mailad1 Or objItem.Sender.Address Like mailad2 Then
    objItem.UnRead = False
    filePath = SAVE_PATH & objItem.Subject & ".txt"
    fileNo = FreeFile
    Open filePath For Append As #fileNo
    Print #fileNo, ""
    Close #fileNo
    
    filePath2 = SAVE_PATH & Mid(objItem.body, InStr(objItem.body, """") + 1, 7) & "識別子.txt"
    fileNo = FreeFile
    Open filePath2 For Append As #fileNo
    Print #fileNo, ""
    Close #fileNo
End If


End Sub
Private Sub payroll(ByVal objItem As MailItem) '給料明細
    Const ATTACH_PATH = "D:\006 給料明細" ' 添付ファイルを保存するフォルダー
    Dim objAttach As attachment
    Dim strFileName As String
    Dim sFN1 As String
    Dim sFN2 As String
    Dim mailad1 As String

    Dim c As Integer
    Dim i As Long
    Dim h1 As Long
    Dim h2 As Long
    Dim objFSO 'As FileSystemObjectDim objOutlook As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    mailad1 = "tanicopayroll@tanico.co.jp"
    'mailad1 = "sn-kumagai@tanico.co.jp"
    
    If objItem.Sender.Address Like mailad1 Then
    objItem.UnRead = False
    
    '添付保存
        For Each objAttach In objItem.Attachments
             c = 1
            With objAttach
                strFileName = .FileName
                strFileName = Right(strFileName, Len(strFileName) - 2)
                sFN1 = Left(strFileName, InStr(1, strFileName, "-") - 1)
                sFN2 = Mid(strFileName, InStr(1, strFileName, "-") + 1, 8)
                strFileName = sFN2 & "-" & sFN1 & ".pdf"
                .SaveAsFile ATTACH_PATH & "\" & strFileName
            End With
        Next
    End If
End Sub



Private Sub SaveAttach(ByVal objItem As MailItem) '工場積算自動保存
    Const ATTACH_PATH = "K:\04 設計\040 House\工場積算結果" ' 添付ファイルを保存するフォルダー
    Dim objAttach As attachment
    Dim strFileName As String
    Dim mailad1 As String
    Dim mailad2 As String
    Dim c As Integer
    Dim i As Long
    Dim h1 As Long
    Dim h2 As Long
    Dim objFSO 'As FileSystemObjectDim objOutlook As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    mailad1 = "k-nakasato@tanico.co.jp"
    'mailad2 = "sn-kumagai@tanico.co.jp"
    
    If objItem.Sender.Address Like mailad1 And _
    objItem.Subject Like "RE: 【積算依頼*" Or _
    objItem.Subject Like "Re: 【積算依頼】*" Then
    
    objItem.UnRead = False
    
    '積算結果保存
        For Each objAttach In objItem.Attachments
             c = 1
            With objAttach
                strFileName = .FileName
                While objFSO.FileExists(ATTACH_PATH & strFileName)
                    strFileName = Left(.FileName, InStrRev(.FileName, ".") - 1) _
                        & "-" & c & Mid(.FileName, InStrRev(.FileName, "."))
                    c = c + 1
                Wend
                h1 = InStr(strFileName, "／")
                h2 = Len(strFileName)
                strFileName = Replace(Mid(strFileName, h1 + 1, h2 - h1), ".pdf", "")
                .SaveAsFile ATTACH_PATH & "\" & strFileName & "_" & Format(Date, "yyyymmdd") & ".pdf"
            End With
        Next
    
        '返信
        If objItem.body Like "*Shunsuke*" Then
        Dim objReply As Object
        Set objReply = objItem.Reply
            With objReply
                .body = _
                    "中里様" & vbCrLf & vbCrLf & _
                    "お疲れ様です。" & vbCrLf & _
                    "積算ありがとうございます。" & vbCrLf & vbCrLf & _
                    "熊谷" & vbCrLf & vbCrLf & .body
                .Save      '下書き保存する
            End With
            objItem.UnRead = True
        End If
    End If
End Sub

Private Sub reply_drawing(ByVal objItem As MailItem) '図面入手返信
    Dim objAttach As attachment
    Dim strFileName As String
    Dim mailad1 As String
    Dim mailad2 As String
    Dim mailad3 As String
    Dim sender_name As String
    Dim c As Integer
    Dim i As Long
    Dim h1 As Long
    Dim h2 As Long
    Dim objFSO 'As FileSystemObjectDim objOutlook As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    mailad1 = "c-makino@nakanishi-mfg.com"
    mailad2 = "k-haraguchi@nakanishi-mfg.com"
    mailad3 = "y-sakashita@nakanishi-mfg.com"
    
    If objItem.Sender.Address Like mailad1 Or _
    objItem.Sender.Address Like mailad2 Or _
    objItem.Sender.Address Like mailad3 Then
    
        If objItem.Sender.Address = mailad1 Then
            sender_name = "牧野様"
        ElseIf objItem.Sender.Address = mailad2 Then
            sender_name = "原口様"
        ElseIf objItem.Sender.Address = mailad3 Then
            sender_name = "坂下様"
        End If
    
        If objItem.Subject Like "RE: 【図面送付依頼*" Or _
        objItem.Subject Like "Re: 【図面送付依頼*" Then
            '返信
            If objItem.body Like "*sn-kumagai@*" Then
            Dim objReply As Object
            Set objReply = objItem.Reply
                With objReply
                    .body = _
                        sender_name & vbCrLf & vbCrLf & _
                        "いつもお世話になっております。" & vbCrLf & _
                        "図面確認いたしました。" & vbCrLf & _
                        "ありがとうございます。" & vbCrLf & vbCrLf & _
                        "**************************" & vbCrLf & _
                        "タニコー株式会社" & vbCrLf & _
                        "機器事業部設計課　熊谷俊介" & vbCrLf & _
                        "Mail:sn-kumagai@tanico.co.jp" & vbCrLf & _
                        "Mobile:070-4948-4042" & vbCrLf & _
                        "〒142-0041" & vbCrLf & _
                        "東京都品川区戸越1-7-20　5F" & vbCrLf & _
                        "**************************" & vbCrLf & _
                        .body
                    .Save      '下書き保存する
                End With
            End If
            objItem.UnRead = True
        End If
    End If
End Sub
'
Private Sub Date_Posting_DDPS(ByVal objItem As MailItem) 'DDPS承認メール受信時転送用メールを作成

Dim mailad As String
Dim String1 As String, String2 As String, String3 As String
Dim STORE_NAME As String
Dim FileName As String
Dim FolderName As String
Dim i As Variant
 
mailad1 = "katsumi.ikoma@jp.mcd.com"
mailad2 = "kazuya.hamakawa@jp.mcd.com" ' "hirokazu.sawada@jp.mcd.com"
 
'件名に"RE: 【レイアウト承認依頼】*"があり"katsumi.ikoma@jp.mcd.com","hirokazu.sawada@jp.mcd.com"から受信した場合
 If objItem.Sender.Address Like mailad1 Or objItem.Sender.Address Like mailad2 Then
 If objItem.Subject Like "RE: 直営PJ【レイアウト承認依頼*D-DPS導入" Then
 If objItem.body Like "*進めてください*" Or objItem.body Like "*承認します*" _
 Or objItem.body Like "*OK*" Then
 
     

'---------転送用メール作成
Dim fwMail As MailItem
Set fwMail = objItem.Forward
Dim Rec As recipient
        
            Set Rec = fwMail.Recipients.add("m-iijima@tanico.co.jp" & "; " & "ru-yamada@tanico.co.jp")
            Rec.Type = olTo
            Set Rec = fwMail.Recipients.add("kenichi-suzuki@tanico.co.jp" & "; " & _
                                            "seino@tanico.co.jp" & "; " & _
                                            "台 健太郎 <k-dai@tanico.co.jp>" & "; " & _
                                            "吉廣 雅幸 <ms-yoshihiro@tanico.co.jp>" & "; " & _
                                            "hr-kurishiro@tanico.co.jp")
            Rec.Type = olCC
            Rec.Resolve
            fwMail.body = _
                "飯島様 山田様" & vbCrLf & vbCrLf & _
                "お疲れ様です。" & vbCrLf & _
                "表題レイアウト店舗承認頂きましたので" & vbCrLf & _
                "シェアポイント、サーバーに登録しました。" & vbCrLf & vbCrLf & _
                "熊谷" & vbCrLf & vbCrLf & fwMail.body
        
        
                  
        fwMail.Save
        objItem.UnRead = False
    End If
    End If
    End If
    
'---------店舗名メモ
Dim BookName As String
Dim app As Object
Dim tname


tname = objItem.Subject
    
    
    
End Sub
'
Private Sub Date_Posting_DDPS_TEST(ByVal objItem As MailItem) 'テストDDPS承認メール受信時転送用メールを作成

Dim mailad As String
Dim String1 As String, String2 As String, String3 As String
Dim STORE_NAME As String
Dim FileName As String
Dim FolderName As String
Dim i As Variant
 
mailad1 = "katsumi.ikoma@jp.mcd.com"
mailad2 = "kazuya.hamakawa@jp.mcd.com" ' "hirokazu.sawada@jp.mcd.com"
 
'件名に"RE: 【レイアウト承認依頼】*"があり"katsumi.ikoma@jp.mcd.com","hirokazu.sawada@jp.mcd.com"から受信した場合
 If objItem.Sender.Address Like mailad1 Or objItem.Sender.Address Like mailad2 Then
 If objItem.Subject Like "RE: 直営PJ【レイアウト承認依頼*D-DPS導入" Then
 If objItem.body Like "*進めてください*" Or objItem.body Like "*承認します*" _
 Or objItem.body Like "*OK*" Then
 
     

'---------転送用メール作成
Dim fwMail As MailItem
Set fwMail = objItem.Forward
Dim Rec As recipient
        
            Set Rec = fwMail.Recipients.add("m-iijima@tanico.co.jp" & "; " & "ru-yamada@tanico.co.jp")
            Rec.Type = olTo
            Set Rec = fwMail.Recipients.add("kenichi-suzuki@tanico.co.jp" & "; " & _
                                            "seino@tanico.co.jp" & "; " & _
                                            "台 健太郎 <k-dai@tanico.co.jp>" & "; " & _
                                            "吉廣 雅幸 <ms-yoshihiro@tanico.co.jp>" & "; " & _
                                            "hr-kurishiro@tanico.co.jp")
            Rec.Type = olCC
            Rec.Resolve
            fwMail.body = _
                "飯島様 山田様" & vbCrLf & vbCrLf & _
                "お疲れ様です。" & vbCrLf & _
                "表題レイアウト店舗承認頂きましたので" & vbCrLf & _
                "シェアポイント、サーバーに登録しました。" & vbCrLf & vbCrLf & _
                "熊谷" & vbCrLf & vbCrLf & fwMail.body
        
        
                  
        fwMail.Save
        objItem.UnRead = False
    End If
    End If
    End If
End Sub

Function LinkCheck(m As MailItem)
    If InStr(1, m.Subject, "IBCテスト") > 0 Then
        With m.GetInspector.WordEditor.HyperLinks
            ReDim arr(1 To .Count)
            Dim i, t, n
            For i = 1 To .Count
                't = .Item(i).TextToDisplay
                n = .Item(i).Name
                'Debug.Print (n)
                If InStr(1, n, "http") > 0 Then
                    MsgBox m.Sender.Address & vbTab & n
                '    arr(i) = Join(Array(n), vbTab)
                End If
            Next
        End With
        'MsgBox Join(arr, vbLf)
    End If
End Function



