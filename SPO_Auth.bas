Attribute VB_Name = "SPO_Auth"
Sub SPO_Auth()
    Dim Driver As New Selenium.WebDriver
    
    Dim SPO_Identifier 'ブラウザ識別子
    Dim sourceFolderPath
    Dim SPO_URL
    Dim MailAdress
    Dim CodeFolder
    Dim AuthVerifiCode 'パスコード
    Dim AuthIdentifier 'メール識別子
    Dim i As Integer

'ダウンロード場所
    sourceFolderPath = "**************" & "\"
'SharePointのURL
    SPO_URL = "**************"
'Outlookに届くメールアドレス
    MailAdress = "**************"
'コードを保存するフォルダ ※CodeFolderはSPO_Auth_Mailと同じフォルダ
    CodeFolder = "**************" & "\"
    
'ブラウザ操作//
    'headlessにすると不安定になるのでコメントアウトしています。
    'Driver.AddArgument "headless"
    Driver.SetPreference "download.default_directory", sourceFolderPath
    Driver.Start "chrome"
    Driver.Window.SetSize 1300, 800
    Driver.Get SPO_URL
    
'ログイン処理ここから
    'メールアドレス記入
    Driver.FindElementByXPath("//*[@id=""txtTOAAEmail""]").SendKeys MailAdress
    'ボタンクリック
    Driver.FindElementByXPath("//*[@id=""btnSubmitEmail""]").Click

    Do
        i = i + 1
        
        If i <> 1 Then
            Driver.FindElementByXPath("//*[@id=""lnkSendCode""]").Click
        End If
        
        Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").SendKeys "1111111" '識別子出すためのダミー
        Driver.FindElementByXPath("//*[@id=""btnSubmitCode""]").Click
        SPO_Identifier = Driver.FindElementByXPath("//*[@id=""ValidateTOAACodeText""]/b").Attribute("innerHTML")
        Call SPO_AuthCode(CodeFolder, AuthVerifiCode, AuthIdentifier)

    Loop Until AuthIdentifier = SPO_Identifier

    Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").Clear
    Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").SendKeys AuthVerifiCode
    Driver.FindElementByXPath("//*[@id=""btnSubmitCode""]").Click
'ログイン処理ここまで
    
'ここから希望の処理
    Call Sample(Driver)
'希望の処理ここまで
    
    Driver.Close
    
End Sub
'ワンタイムコードと識別子を抽出
Sub SPO_AuthCode(ByRef CodeFolder, ByRef AuthVerifiCode, ByRef AuthIdentifier)
    Dim startTime As Double
    Dim endTime As Double
    Dim TimeDifference As Double
    Dim CountF
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    CountF = fso.GetFolder(CodeFolder).Files.Count
    '既存テキストファイル削除
    If CountF > 0 Then
        Kill CodeFolder & "*.txt"
        CountF = 0
    End If
    
    'OUTLOOK操作
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olInbox As Outlook.MAPIFolder
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    'OUTLOOKが立ち上がってなかったら立ち上げる
    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
        Set olNs = olApp.GetNamespace("MAPI")
        Set olInbox = olNs.GetDefaultFolder(olFolderInbox)
        olInbox.Display
    End If
    
    Call olApp.Session.LogOn("Outlook", "")

    'メール送受信してフォルダに保存されるまでループ
    startTime = Timer
    Do
        Call olApp.Session.SendAndReceive(True) '全て送受信
        Application.Wait Now() + TimeValue("00:00:05")
        CountF = fso.GetFolder(CodeFolder).Files.Count
        endTime = Timer
        TimeDifference = endTime - startTime
        Debug.Print TimeDifference
        
        '300秒たったら抜けて再度コード発行
        If TimeDifference > 300 Then
            Exit Sub
        End If
    Loop Until CountF > 0
    
    
    '保存されたtxtファイルからcodeを抽出
    buf = Dir(CodeFolder & "*.txt")
    Do While buf <> ""
        If InStr(buf, "検証コード") > 0 Then
        AuthVerifiCode = Left(buf, 8)
        End If
        
        If InStr(buf, "識別") > 0 Then
        AuthIdentifier = Left(buf, 7)
        End If
        
        buf = Dir()
    Loop
    
    Set fso = Nothing
    Set olApp = Nothing
    Set olNs = Nothing
    Set olInbox = Nothing
End Sub

Function Sample(Driver)

Driver.Get "https://google.com"

End Function
