Attribute VB_Name = "Dictionary"

'***********************************************************************************************************************
' �@�\   : �����o�^�@�\
' �T�v   : ���͂��ꂽ�����A�o�^���������擾���āA�G�N�Z���ɕ\������
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Function SearchFromDic()

    On Error GoTo errorHandler

    'D-Tools��ʂ̏����ݒ�����{����
    Load DTools

    Dim selectSQL As String
    selectSQL = "SELECT [�_����],[������],[���l],[�ǉ���],[�ǉ���],[�폜�t���O] FROM [�_���ϊ��e�[�u��] "
    Dim selectSQLConditions  As String
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    For i = 2 To rowCount
        Dim selectSQLCondition  As String
        selectSQLCondition = ""
        If ActiveSheet.Cells(i, 1).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[�_����] like '%" & ActiveSheet.Cells(i, 1).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 2).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[������] like '%" & ActiveSheet.Cells(i, 2).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 3).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[���l] like '%" & ActiveSheet.Cells(i, 3).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 4).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[�ǉ���] like '%" & ActiveSheet.Cells(i, 4).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 5).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[�ǉ���] like '%" & ActiveSheet.Cells(i, 5).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 6).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[�폜�t���O] like '%" & ActiveSheet.Cells(i, 6).value & "%' AND "
        End If

        If selectSQLCondition <> "" Then
            selectSQLCondition = Mid(selectSQLCondition, 1, Len(selectSQLCondition) - 4)
            selectSQLCondition = "(" & selectSQLCondition & ") OR "
            
            selectSQLConditions = selectSQLConditions & selectSQLCondition
        End If
        
    Next
    
    If selectSQLConditions <> "" Then
        selectSQLConditions = Mid(selectSQLConditions, 1, Len(selectSQLConditions) - 3)
        selectSQL = selectSQL & " WHERE " & selectSQLConditions
    End If
    
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Set ADOConnection = connAccessDB()
    '����SQL�����s����
    ADORecordset.Open selectSQL, ADOConnection
    
    i = 2
    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        ActiveSheet.Cells(i, 1).value = resultFields("�_����").value
        ActiveSheet.Cells(i, 2).value = resultFields("������").value
        ActiveSheet.Cells(i, 3).value = resultFields("���l").value
        ActiveSheet.Cells(i, 4).value = resultFields("�ǉ���").value
        ActiveSheet.Cells(i, 5).value = resultFields("�ǉ���").value
        ActiveSheet.Cells(i, 6).value = resultFields("�폜�t���O").value
        i = i + 1
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    'G����N���A����
    ActiveSheet.columns("G").Clear
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("SearchFromDic")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If
    
End Function

'***********************************************************************************************************************
' �@�\   : �����o�^�@�\
' �T�v   :  ���͂��ꂽ����_�������ɓo�^����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Function RegisterToDic()

On Error GoTo errorHandler

    'D-Tools��ʂ̏����ݒ�����{����
    Load DTools
    
    Dim selectSQL As String
    Dim updateSQL As String
    Dim insertSQL As String
    Dim deleteSQL As String
    Dim sqlConditions  As String
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Set ADOConnection = connAccessDB()
    
    'G����N���A����
    ActiveSheet.columns("G").Clear
    
    For i = 2 To rowCount
           
            �_���� = ActiveSheet.Cells(i, 1).value
            ������ = ActiveSheet.Cells(i, 2).value
            ���l = ActiveSheet.Cells(i, 3).value
            �ǉ��� = ActiveSheet.Cells(i, 4).value
            �ǉ��� = ActiveSheet.Cells(i, 5).value
            �폜�t���O = ActiveSheet.Cells(i, 6).value
            
            If �_���� <> "" And ������ <> "" And �ǉ��� <> "" And �ǉ��� <> "" And �폜�t���O <> "" Then

                selectSQL = "SELECT * FROM [�_���ϊ��e�[�u��] WHERE [�폜�t���O] = '0' AND ([�_����] = '" & �_���� & "' AND [������] = '" & ������ & "')"
                
                '�����̃`�F�b�N
                Set ADORecordset = New ADODB.recordset
                ADORecordset.Open selectSQL, ADOConnection
                
                '���������ԂŁA�V�K�o�^�̏ꍇ
                If �폜�t���O = "0" And ADORecordset.EOF = False Then
                    ActiveSheet.Cells(i, 7).value = "�_�����ƕ����������ɓo�^����܂����B"
                    ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                End If
                
                '�V�K�o�^�̏ꍇ
                If �폜�t���O = "0" And ADORecordset.EOF = True Then
                    '�����̎������A�ϊ��ł��邩�ǂ����`�F�b�N����
                    'ConvertStrInLoop (�_����)
                    insertSQL = "INSERT INTO [�_���ϊ��e�[�u��] ([�_����],[������],[���l],[�ǉ���],[�ǉ���],[�폜�t���O]) VALUES ('" & �_���� & "','" & ������ & "','" & ���l & "','" & �ǉ��� & "','" & �ǉ��� & "','" & �폜�t���O & "')"
                    ADOConnection.Execute (insertSQL)
                    ActiveSheet.Cells(i, 7).value = "�o�^�ς�"
                End If
                
                
                If �폜�t���O = "1" Then
                    selectSQL = "SELECT * FROM [�_���ϊ��e�[�u��] WHERE [�_����] = '" & �_���� & "' AND [������] = '" & ������ & "'"
                    Set ADORecordset = New ADODB.recordset
                    ADORecordset.Open selectSQL, ADOConnection
                    '�����ɑ΂��Ę_���폜����ꍇ
                    If ADORecordset.EOF = False Then
                        updateSQL = "UPDATE [�_���ϊ��e�[�u��] SET [���l] = '" & ���l & "',[�ǉ���] = '" & �ǉ��� & "',[�ǉ���]= '" & �ǉ��� & "',[�폜�t���O] = '" & �폜�t���O & "' WHERE [�폜�t���O] = '0' AND ([�_����] = '" & �_���� & "' AND [������] = '" & ������ & "')"
                        ADOConnection.Execute (updateSQL)
                        ActiveSheet.Cells(i, 7).value = "�_���폜�X�V�ς�"
                    Else
                        ActiveSheet.Cells(i, 7).value = "�����폜���ł��܂���B�Ώۂ����|�W�g���ɑ��݂��Ȃ��B"
                        ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                    End If
                End If
                
                If �폜�t���O <> "0" And �폜�t���O <> "1" Then
                    selectSQL = "SELECT * FROM [�_���ϊ��e�[�u��] WHERE [�_����] = '" & �_���� & "' AND [������] = '" & ������ & "'"
                    Set ADORecordset = New ADODB.recordset
                    ADORecordset.Open selectSQL, ADOConnection
                    '�����ɑ΂��ĕ����폜����ꍇ
                    If ADORecordset.EOF = False Then
                        deleteSQL = "DELETE FROM [�_���ϊ��e�[�u��] WHERE [�_����] = '" & �_���� & "' AND [������] = '" & ������ & "'"
                        ADOConnection.Execute (deleteSQL)
                        ActiveSheet.Cells(i, 7).value = "�����폜�ς�"
                    Else
                        ActiveSheet.Cells(i, 7).value = "�����폜���ł��܂���B�Ώۂ����|�W�g���ɑ��݂��Ȃ��B"
                        ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                    End If
                    
                End If

            Else
                ActiveSheet.Cells(i, 7).value = "�o�^���s�I�_�����A�������A�ǉ��ҁA�ǉ����A�폜�t���O���K�v�̂��߁A�ݒ肵�Ă��������B"
                ActiveSheet.Cells(i, 7).Font.colorIndex = 3
            End If

    Next
        
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("RegisterToDic")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If
       
    ADORecordset.Close
    ADOConnection.Close
    
End Function

'***********************************************************************************************************************
' �@�\   : �_���ϊ��@�\
' �T�v   : ���͂��ꂽ�_�����𕨗����֕ϊ�����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Function LogicalToPhysicalByDic()
    On Error GoTo errorHandler
    
    'D-Tools��ʂ̏����ݒ�����{����
    Load DTools
        
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    '�S�̗p�ϐ�������������
    PUB_TEMP_VAR_STR = ""
    
    For i = 2 To rowCount
        �_���� = ActiveSheet.Cells(i, 1).value
        ActiveSheet.Cells(i, 2).value = ""
        ActiveSheet.Cells(i, 3).value = ""
        
        
        '�_���ϊ����\�b�h���Ăт���
        ConvertStrInLoop (�_����)
        
        '�ϊ��������e���G�N�Z���ɏo�͂���
        ActiveSheet.Cells(i, 2).value = PUB_TEMP_VAR_STR
        
        If PUB_TEMP_VAR_STR <> "" Then
            '�ϊ��ł��Ȃ������i�p���ȂǈȊO�̕����j���ĕϊ�����
            PUB_TEMP_VAR_STR = ConvertHiraganaToEnglish(PUB_TEMP_VAR_STR)
            If ActiveSheet.CheckBoxes.value = xlOn Then
                If ActiveSheet.Cells(i, 2).value <> PUB_TEMP_VAR_STR Then
                    ActiveSheet.Cells(i, 3).value = "�_���������ϊ��ł��Ȃ������𑶍݂��Ă��܂��B" & vbCrLf & "�y" & PUB_TEMP_VAR_STR & "�z��_�������ɓo�^���܂���"
                End If
            Else
                If ActiveSheet.Cells(i, 2).value <> PUB_TEMP_VAR_STR Then
                    ActiveSheet.Cells(i, 3).value = "�_���������ϊ��ł��Ȃ������𑶍݂��Ă��܂��B"
                End If
            End If
            
        End If
        
        '�S�̗p�ϐ�������������
        PUB_TEMP_VAR_STR = ""
    Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("LogicalToPhysicalByDic")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If
    
End Function

'***********************************************************************************************************************
' �@�\   : �_���ϊ��@�\
' �T�v   : ���͂��ꂽ�������Loop���āA�������ɕϊ�����B
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Function ConvertStrInLoop(convertStr As String)
    Dim RecursionFlag As Boolean
    
    For i = 0 To Len(convertStr) - 1
        Dim subConvertStr As String
        Dim subConvertedStr As String
        subConvertStr = Mid(convertStr, 1, Len(convertStr) - i)
        subConvertedStr = ConvertJPToEnglish(subConvertStr)
        
        '��ڂ̕������_���������ϊ��ł��Ȃ��ꍇ�A���}�X�^�ɒ�`������ꍇ
        If Len(subConvertStr) = 1 And ConvertHiraganaToEnglishByMastTab(subConvertStr) <> "" And subConvertedStr = "" Then
            '�ύX�O�̕����̂܂܂�ݒ肷��
            subConvertedStr = subConvertStr
        End If
        
        '�ϊ����ʂ��A�ċA�ϊ�����
        If subConvertedStr <> "" Then
            '�ϊ��ł���������S�̗p�ꎞ�ϐ��ɐݒ肷��
            PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & Replace(Replace(subConvertedStr, " ", ""), "�@", "")
            '�ϊ��ł��������������āA�ȊO�̕�����ύX�ΏۂƂ���B
            subConvertStr = Right(convertStr, Len(convertStr) - Len(subConvertStr))
            If subConvertStr <> "" Then
                ConvertStrInLoop (subConvertStr)
            End If
            Exit For
        End If
        
        '��ڂ̕������ϊ��ł��Ȃ��ꍇ�A�ċA�ϊ�����
        If subConvertedStr = "" And Len(subConvertStr) = 1 Then
            '�}�X�^���A�ύX����B
            'PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & ConvertHiraganaToEnglish(subConvertStr)
            '�ϊ��ł��Ȃ��������S�̗p�ꎞ�ϐ��ɐݒ肷��
            PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & subConvertStr
            '��ڂ̕����������āA�ȊO�̕�����ύX�ΏۂƂ���B
            subConvertStr = Right(convertStr, Len(convertStr) - Len(subConvertStr))
            If subConvertStr <> "" Then
                ConvertStrInLoop (subConvertStr)
            End If
            Exit For
        End If
    Next

End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ���{��P��̕��������畨�����ɕϊ�����
' ����   : String ������
' �߂�l : String �������̕�����
'***********************************************************************************************************************
Private Function ConvertHiraganaToEnglish(convertStr As String) As String
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Dim ������ As String
    Dim convertStrToHiragana As String
    Set ADOConnection = connAccessDB()
    
    For j = 1 To Len(convertStr)
        
        If Mid(convertStr, j, 1) Like "*[a-z,A-Z,0-9,_]" Then
            ������ = ������ & Mid(convertStr, j, 1)
        Else
            convertStrToHiragana = StrConv(Application.GetPhonetic(Mid(convertStr, j, 1)), vbHiragana)
            For i = 1 To Len(convertStrToHiragana)
                Dim hi As String
                Dim convertHIToEN As String
                hi = Mid(convertStrToHiragana, i, 1)
                
                convertHIToEN = ConvertHiraganaToEnglishByMastTab(hi)
                
                ������ = ������ & ConvertHiraganaToEnglishByMastTab(hi)
                
            Next
        End If
    Next
    
    ConvertHiraganaToEnglish = ������
    
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ���{��P��𕽉����p���}�b�s���O�}�X�^��蕨�����ɕϊ�����
' ����   : String�@���{��P��
' �߂�l : String�@������
'***********************************************************************************************************************
Private Function ConvertHiraganaToEnglishByMastTab(convertStr As String) As String
    Dim ������ As String
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    
    Dim selectSQL As String
    selectSQL = "SELECT [������] FROM [�������p���}�b�s���O�}�X�^] WHERE [�폜�t���O] = '0' AND [�_����] = '" & convertStr & "'"
    
    '�����̃`�F�b�N
    Dim ADORecordset As New ADODB.recordset
    ADORecordset.Open selectSQL, ADOConnection
    
    Do Until ADORecordset.EOF
        Dim resultFields As Fields
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("������").value) Then
            ������ = ""
        Else
            ������ = resultFields("������").value
        End If
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    ConvertHiraganaToEnglishByMastTab = ������
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �_�������畨�����ɕϊ�����
' ����   : String�@�_����
' �߂�l : String�@������
'***********************************************************************************************************************
Private Function ConvertJPToEnglish(convertStr As String) As String
    Dim ������ As String
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    
    Dim selectSQL As String
    selectSQL = "SELECT [������] FROM [�_���ϊ��e�[�u��] WHERE [�폜�t���O] = '0' AND [�_����] = '" & convertStr & "'"
    
    '�����̃`�F�b�N
    Dim ADORecordset As New ADODB.recordset
    ADORecordset.Open selectSQL, ADOConnection
    
    Do Until ADORecordset.EOF
        Dim resultFields As Fields
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("������").value) Then
            ������ = ""
        Else
            ������ = resultFields("������").value
        End If
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    ConvertJPToEnglish = ������
    
End Function
