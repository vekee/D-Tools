Attribute VB_Name = "UpVersion"
Public Const ADDIN_FILE_NAME = "D-Tools.xlam"
Public Const UP_VERION_SCRIPT_FILE_NAME = "D-ToolsUpVersion.vbs"

'***********************************************************************************************************************
' �@�\   : �c�[���X�V�X�N���v�g�o�͋@�\
' �T�v   : �G�N�Z�����J���������A�c�[���o�[�ŁA�{�^�����쐬����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Function CreateUpVersionScript()
    Dim adoObject As Object
    Set adoObject = CreateObject("ADODB.Stream")
    adoObject.Type = adTypeText
    adoObject.Charset = "SJIS"
    adoObject.LineSeparator = adCRLF
    adoObject.Open
    
    adoObject.WriteText ("Option Explicit" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("Dim objFileSys" & vbCrLf)
    adoObject.WriteText ("Dim strFilePathFrom" & vbCrLf)
    adoObject.WriteText ("Dim strFilePathTo" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'3�b�X���[�v" & vbCrLf)
    adoObject.WriteText ("call WScript.Sleep(3*1000)" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'�t�@�C���V�X�e���������I�u�W�F�N�g���쐬" & vbCrLf)
    adoObject.WriteText ("Set objFileSys = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'�R�s�[���̃t�@�C���̃p�X���w��" & vbCrLf)
    adoObject.WriteText ("strFilePathFrom = ""\\cs.dir.co.jp\document\���ʃG���A\HN_Project\PJ_BSNS-20121227-00005_M��G�V��i�p�[�g�i���L�j\30.ZETTA���A���^�C�����E�N���E�h��\99.WORK\90.JS\98_�c�[��\D-Tools\" & ADDIN_FILE_NAME & """" & vbCrLf)
    adoObject.WriteText ("strFilePathTo   = " & """" & GetAddinDir & "\" & ADDIN_FILE_NAME & """")
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'�G���[�������ɂ������𑱍s����悤�ݒ�" & vbCrLf)
    adoObject.WriteText ("On Error Resume Next" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'�t�@�C�����㏑���R�s�[" & vbCrLf)
    adoObject.WriteText ("Call objFileSys.CopyFile(strFilePathFrom, strFilePathTo)" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'���t�@�C���������ȂǁA�G���[�ɂȂ����ꍇ�̏���" & vbCrLf)
    adoObject.WriteText ("If Err.Number <> 0 Then" & vbCrLf)
    adoObject.WriteText ("  '�������Ȃ�" & vbCrLf)
    adoObject.WriteText ("  '�G���[�����N���A����B" & vbCrLf)
    adoObject.WriteText ("  Err.Clear" & vbCrLf)
    adoObject.WriteText ("End If" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'�uOn Error Resume Next�v������" & vbCrLf)
    adoObject.WriteText ("On Error Goto 0" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("Set objFileSys = Nothing " & vbCrLf)
    
    
    Dim vbsFile As String
    vbsFile = Application.UserLibraryPath & "\" & UP_VERION_SCRIPT_FILE_NAME
    
       '�t�@�C����ۑ�����
    adoObject.SaveToFile (vbsFile), adSaveCreateOverWrite
    '�t�@�C���ƕ���
    adoObject.Close
    
End Function
'***********************************************************************************************************************
' �@�\   : �G�N�Z���A�h�C���̃p�X�擾
' �T�v   :�G�N�Z���A�h�C���̃p�X���擾����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Function GetAddinDir() As String
    GetAddinDir = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Roaming\Microsoft\AddIns"
End Function
