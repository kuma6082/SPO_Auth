Attribute VB_Name = "SPO_Auth"
Sub SPO_Auth()
    Dim Driver As New Selenium.WebDriver
    
    Dim SPO_Identifier '�u���E�U���ʎq
    Dim sourceFolderPath
    Dim SPO_URL
    Dim MailAdress
    Dim CodeFolder
    Dim AuthVerifiCode '�p�X�R�[�h
    Dim AuthIdentifier '���[�����ʎq
    Dim i As Integer

'�_�E�����[�h�ꏊ
    sourceFolderPath = "**************" & "\"
'SharePoint��URL
    SPO_URL = "**************"
'Outlook�ɓ͂����[���A�h���X
    MailAdress = "**************"
'�R�[�h��ۑ�����t�H���_
    CodeFolder = "**************" & "\"
    
'�u���E�U����//
    'headless�ɂ���ƕs����ɂȂ�̂ŃR�����g�A�E�g���Ă��܂��B
    'Driver.AddArgument "headless"
    Driver.SetPreference "download.default_directory", sourceFolderPath
    Driver.Start "chrome"
    Driver.Window.SetSize 1300, 800
    Driver.Get SPO_URL
    
'���O�C��������������
    '���[���A�h���X�L��
    Driver.FindElementByXPath("//*[@id=""txtTOAAEmail""]").SendKeys MailAdress
    '�{�^���N���b�N
    Driver.FindElementByXPath("//*[@id=""btnSubmitEmail""]").Click

    Do
        i = i + 1
        
        If i <> 1 Then
            Driver.FindElementByXPath("//*[@id=""lnkSendCode""]").Click
        End If
        
        Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").SendKeys "1111111" '���ʎq�o�����߂̃_�~�[
        Driver.FindElementByXPath("//*[@id=""btnSubmitCode""]").Click
        SPO_Identifier = Driver.FindElementByXPath("//*[@id=""ValidateTOAACodeText""]/b").Attribute("innerHTML")
        Call SPO_AuthCode(CodeFolder, AuthVerifiCode, AuthIdentifier)

    Loop Until AuthIdentifier = SPO_Identifier

    Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").Clear
    Driver.FindElementByXPath("//*[@id=""txtTOAACode""]").SendKeys AuthVerifiCode
    Driver.FindElementByXPath("//*[@id=""btnSubmitCode""]").Click
'���O�C�����������܂�
    
'���������]�̏���
    Call Sample(Driver)
'��]�̏��������܂�
    
    Driver.Close
    
End Sub
'�����^�C���R�[�h�Ǝ��ʎq�𒊏o
Sub SPO_AuthCode(ByRef CodeFolder, ByRef AuthVerifiCode, ByRef AuthIdentifier)
    Dim startTime As Double
    Dim endTime As Double
    Dim TimeDifference As Double
    Dim CountF
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    CountF = fso.GetFolder(CodeFolder).Files.Count
    '�����e�L�X�g�t�@�C���폜
    If CountF > 0 Then
        Kill CodeFolder & "*.txt"
        CountF = 0
    End If
    
    'OUTLOOK����
    Dim olApp As Outlook.Application
    Dim olNs As Outlook.Namespace
    Dim olInbox As Outlook.MAPIFolder
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    On Error GoTo 0
    
    'OUTLOOK�������オ���ĂȂ������痧���グ��
    If olApp Is Nothing Then
        Set olApp = New Outlook.Application
        Set olNs = olApp.GetNamespace("MAPI")
        Set olInbox = olNs.GetDefaultFolder(olFolderInbox)
        olInbox.Display
    End If
    
    Call olApp.Session.LogOn("Outlook", "")

    '���[������M���ăt�H���_�ɕۑ������܂Ń��[�v
    startTime = Timer
    Do
        Call olApp.Session.SendAndReceive(True) '�S�đ���M
        Application.Wait Now() + TimeValue("00:00:05")
        CountF = fso.GetFolder(CodeFolder).Files.Count
        endTime = Timer
        TimeDifference = endTime - startTime
        Debug.Print TimeDifference
        
        '300�b�������甲���čēx�R�[�h���s
        If TimeDifference > 300 Then
            Exit Sub
        End If
    Loop Until CountF > 0
    
    
    '�ۑ����ꂽtxt�t�@�C������code�𒊏o
    buf = Dir(CodeFolder & "*.txt")
    Do While buf <> ""
        If InStr(buf, "���؃R�[�h") > 0 Then
        AuthVerifiCode = Left(buf, 8)
        End If
        
        If InStr(buf, "����") > 0 Then
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
