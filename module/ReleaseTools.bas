Attribute VB_Name = "ReleaseTools"
'���ރ����[�X�p�c�[��
Sub ExtractFiles()
    
    '�S�̒萔�𗘗p����
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    

    '�o�̓t�H���_�[
    outputRootFolder = GetSaveDir & "\" & "���ޒ��o����_" & Format(Now, "yyyymmddHHMMSS")

    For i = 2 To maxRowNo
        
        On Error GoTo errorHandler
        '���[�gDIR
        rootDir = normalizePath(ActiveSheet.Cells(i, 1).value)
        '�t�@�C���p�X
        filePath = normalizePath(ActiveSheet.Cells(i, 2).value)
        
        '�t�@�C���̑S�p�X
        fileAllPath = rootDir & "\" & filePath
        
        
        '�T�u�t�H���_�[�쐬
        subFolder = Left(filePath, InStrRev(filePath, "\"))
        MkDir outputRootFolder & "\" & subFolder

        
        '�Ώێ��ނ��o�̓t�H���_�[�ɃR�s�[����
        FileCopy fileAllPath, outputRootFolder & "\" & filePath
        
        ActiveSheet.Cells(i, 3).value = "OK"
        
errorHandler:
    If Err.Number <> 0 Then
        ActiveSheet.Cells(i, 3).value = "NG" & "�F" & Err.Description
    End If

    Next
    
    MsgBox "���o����" & vbCrLf & outputRootFolder
    
End Sub
 
Private Function normalizePath(paramPath As String) As String
               
        paramPath = Replace(paramPath, "/", "\")
        
        If Left(paramPath, 1) = "\" Then
            paramPath = Right(paramPath, Len(paramPath) - 1)
        End If
        
        If Right(paramPath, 1) = "\" Then
            paramPath = Left(paramPath, Len(paramPath) - 1)
        End If
        
        normalizePath = paramPath
        
End Function
 
 
