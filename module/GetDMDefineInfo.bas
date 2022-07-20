Attribute VB_Name = "GetDMDefineInfo"
'***********************************************************************************************************************
' �@�\   : DM��`�������@�\
' �T�v   : �w��f�B���N�g�z����DM���ʕ����ADM��`�f�[�^�X�[�X�t�@�C���𓯊�����
' ����   : Folder DM���ʕ��̃f�B���N�g���AString �����Ώۂ̃f�[�^�\�[�X�t�@�C���̃p�X
' �߂�l : True : ��������   False : �������s
'***********************************************************************************************************************
Public Function SynchronizeDMDefineInfo(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    
    
    tableDefineSheetName = "�e�[�u����`��"
    
    'On Error GoTo errorHandler
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
        

    '�A�E�g�i���o�[������������
    ADOConnection.Execute ("ALTER TABLE [������`��] ALTER COLUMN [ID] COUNTER (1,1)")
    '�A�E�g�i���o�[������������
    ADOConnection.Execute ("ALTER TABLE [�e�[�u����`��] ALTER COLUMN [ID] COUNTER (1,1)")
    
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
        If dirFolderFile.name Like "*�e�[�u����`��*.xls*" Then
        
            For Each objSheet In dirWorkbook.Worksheets
                If objSheet.name <> "�L�����[��" And objSheet.name <> "�ύX����" And objSheet.name <> "�ڎ�" And objSheet.name <> "�V�[�P���X��`" And objSheet.name <> "�ʎ�" Then
                    
                    tableDefineSheetName = objSheet.name
                    
                    Dim �Ő� As String
                    Dim �C���� As String
                    Dim �C���� As String
                    Dim �e�[�u����_�a�� As String
                    Dim �e�[�u����_�p�� As String
                    Dim �e�[�u������ As String
                    Dim insertSql_table As String
                    Dim i As Integer
                    
                    insertSql_table = "INSERT INTO [�e�[�u����`��] ([�e�[�u����_�a��],[�e�[�u����_�p��],[�e�[�u������],[�Ő�],[�C����],[�C����]) VALUES "
                    
        
                    �Ő� = ""
                    �C���� = ""
                    �C���� = ""
                    
        
                    �e�[�u����_�a�� = dirWorkbook.Worksheets(tableDefineSheetName).Range("C3").value
                    �e�[�u����_�p�� = dirWorkbook.Worksheets(tableDefineSheetName).Range("E3").value
                    �e�[�u������ = ""
                    
                    insertSql_table = insertSql_table & "('" & �e�[�u����_�a�� & "'," _
                                                      & "'" & �e�[�u����_�p�� & "'," _
                                                      & "'" & �e�[�u������ & "'," _
                                                      & "'" & �Ő� & "'," _
                                                      & "'" & �C���� & "'," _
                                                      & "'" & �C���� & "')"
                    '�e�[�u����`���̃e�[�u�����N���A����
                    ADOConnection.Execute ("DELETE FROM [�e�[�u����`��] WHERE [�e�[�u����_�p��] = '" & �e�[�u����_�p�� & "'")
                    '������`���̃e�[�u�����N���A����
                    ADOConnection.Execute ("DELETE FROM [������`��] WHERE [�e�[�u����_�p��] = '" & �e�[�u����_�p�� & "'")
                    
                    
                    '�e�[�u����`���̃e�[�u���֔��f����
                    ADOConnection.Execute (insertSql_table)
                    
        
                    i = 8
                    Do While dirWorkbook.Worksheets(tableDefineSheetName).Cells(i, 4) <> "" And dirWorkbook.Worksheets(tableDefineSheetName).Cells(i, 5) <> ""
                        Dim No As String
                        Dim ������_�a�� As String
                        Dim �J������_�p�� As String
                        Dim ��L�[ As String
                        Dim NullAble As String
                        Dim �f�[�^�^ As String
                        Dim ���� As String
                        Dim �����ȉ����� As String
                        Dim �f�B�t�H���g�l As String
                        Dim ��������_�a�� As String
                        Dim insertSql_columns As String
                        
                        insertSql_columns = "INSERT INTO [������`��] ([�e�[�u����_�p��],[NO],[������_�a��],[�J������_�p��],[��L�[],[NULL],[�f�[�^�^],[����],[�����ȉ�����],[�f�B�t�H���g�l],[��������_�a��],[�Ő�],[�C����],[�C����]) VALUES "
                        
                        
                        No = dirWorkbook.Worksheets(tableDefineSheetName).Range("C" & i).value
                        ������_�a�� = dirWorkbook.Worksheets(tableDefineSheetName).Range("D" & i).value
                        �J������_�p�� = dirWorkbook.Worksheets(tableDefineSheetName).Range("E" & i).value
                        ��L�[ = dirWorkbook.Worksheets(tableDefineSheetName).Range("J" & i).value
                        ��L�[ = Replace(��L�[, " ", "")
                        ��L�[ = Replace(��L�[, "�@", "")
                        NullAble = dirWorkbook.Worksheets(tableDefineSheetName).Range("K" & i).value
                        �f�[�^�^ = dirWorkbook.Worksheets(tableDefineSheetName).Range("F" & i).value
                        ���� = dirWorkbook.Worksheets(tableDefineSheetName).Range("G" & i).value
                        �����ȉ����� = dirWorkbook.Worksheets(tableDefineSheetName).Range("H" & i).value
                        �f�B�t�H���g�l = Replace(dirWorkbook.Worksheets(tableDefineSheetName).Range("L" & i).value, "'", "''")
                        ��������_�a�� = ""
                        
                        insertSql_columns = insertSql_columns & "('" & �e�[�u����_�p�� & "'," _
                                                              & Val(No) & "," _
                                                              & "'" & ������_�a�� & "'," _
                                                              & "'" & �J������_�p�� & "'," _
                                                              & "'" & ��L�[ & "'," _
                                                              & "'" & NullAble & "'," _
                                                              & "'" & �f�[�^�^ & "'," _
                                                              & Val(����) & "," _
                                                              & Val(�����ȉ�����) & "," _
                                                              & "'" & �f�B�t�H���g�l & "'," _
                                                              & "'" & ��������_�a�� & "'," _
                                                              & "'" & �Ő� & "'," _
                                                              & "'" & �C���� & "'," _
                                                              & "'" & �C���� & "')"
        
                        '������`���̃e�[�u���֔��f����
                        ADOConnection.Execute (insertSql_columns)
                        
                        i = i + 1
                    Loop
                
                End If
            Next
        
            
        End If
            
        '���������G�N�Z�����N���b�Y����B
        dirWorkbook.Close (False)
        
    Next
    
    '�A�E�g�i���o�[������������
    ADOConnection.Execute ("ALTER TABLE [������`��] ALTER COLUMN [ID] COUNTER (1,1)")
    '�A�E�g�i���o�[������������
    ADOConnection.Execute ("ALTER TABLE [�e�[�u����`��] ALTER COLUMN [ID] COUNTER (1,1)")
    
    '�ċA���������Ȃ�
    'For Each subFolder In dirFolder.SubFolders
    '    Call SearchExcleByContentFromDir(subFolder, searchByContent)
    'Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("SynchronizeDMDefineInfo")
        result = False
    End If
    
    ADOConnection.Close
    
    SynchronizeDMDefineInfo = result
    
End Function
'***********************************************************************************************************************
' �@�\   : �e�[�u����`���W��@�\
' �T�v   : �w��f�B���N�g�z����DM���ʕ����A��G�N�Z���ɏW�񂷂�@�\
' ����   : Folder DM���ʕ��̃f�B���N�g���AString �����Ώۂ̃f�[�^�\�[�X�t�@�C���̃p�X
' �߂�l : True : ��������   False : �������s
'***********************************************************************************************************************
Public Function ExcleExtracte(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    fileCounter = 1
    
    tableListSheetName = "�ڎ�"
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                If dirWorksheet.name <> "���ڐ���" And Not (dirWorksheet.name Like "*�L����*") Then
                    
                    '�V�[�g��
                    tableDefineSheetName = dirWorksheet.name
                    
                    '�ڎ�
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = dirWorksheet.Cells(3, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirWorksheet.Cells(4, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 4).value = dirWorksheet.Cells(2, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 5).value = dirFolderFile.name
        
                    '�e�[�u���V�[�g�쐬
                    Call createNewSheet(tableDefineSheetName, OPERATION_WORKBOOK)
                    
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 1) = "����"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 2) = "���ږ���"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 3) = "�K�w"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 4) = "������"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 5) = "���"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 6) = "�o�C�g��"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 7) = "����"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 8) = "����"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 9) = "�J�n�ʒu"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 10) = "�I���ʒu"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 11) = "����"
                    
                    '�e�e�[�u���V�[�g
                    i = 8
                    Do While dirWorksheet.Cells(i, 2) <> "" Or dirWorksheet.Cells(i, 14) <> "" Or dirWorksheet.Cells(i, 15) <> "" Or dirWorksheet.Cells(i, 29) <> ""
                       
                        
                        ���� = dirWorksheet.Range("A" & i).value
                        ���ږ��� = dirWorksheet.Range("B" & i).value _
                                 & dirWorksheet.Range("C" & i).value _
                                 & dirWorksheet.Range("D" & i).value _
                                 & dirWorksheet.Range("E" & i).value _
                                 & dirWorksheet.Range("F" & i).value _
                                 & dirWorksheet.Range("G" & i).value _
                                 & dirWorksheet.Range("H" & i).value _
                                 & dirWorksheet.Range("I" & i).value _
                                 & dirWorksheet.Range("J" & i).value _
                                 & dirWorksheet.Range("K" & i).value _
                                 & dirWorksheet.Range("L" & i).value _
                                 & dirWorksheet.Range("M" & i).value
                                 
                        �K�w = dirWorksheet.Range("N" & i).value
                        ������ = dirWorksheet.Range("O" & i).value
                        ��� = dirWorksheet.Range("Z" & i).value
                        �o�C�g�� = dirWorksheet.Range("AC" & i).value & dirWorksheet.Range("AD" & i).value
                        If ��� = "P" Then
                            ���� = dirWorksheet.Range("AE" & i).value & "." & dirWorksheet.Range("AF" & i).value & dirWorksheet.Range("AG" & i).value
                        End If
                        ���� = dirWorksheet.Range("AI" & i).value
                        �J�n�ʒu = dirWorksheet.Range("AJ" & i).value
                        �I���ʒu = dirWorksheet.Range("AK" & i).value
                        ���� = dirWorksheet.Range("AM" & i).value
        
        
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 1) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 2) = ���ږ���
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 3) = �K�w
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 4) = ������
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 5) = ���
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 6) = �o�C�g��
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 7) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 8) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 9) = �J�n�ʒu
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 10) = �I���ʒu
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 11) = ����
                            
                        
                        i = i + 1
                    Loop
                    
                    fileCounter = fileCounter + 1
                    
                    
                    '�r����t����
                     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 5).Address).Borders.LineStyle = xlContinuous
                    '��̕���������
                     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 5).Address).columns.AutoFit
                     
                    '�r����t����
                     OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Range(Cells(1, 1).Address & ":" & Cells(i - 7, 11).Address).Borders.LineStyle = xlContinuous
                    '��̕���������
                     OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Range(Cells(1, 1).Address & ":" & Cells(i - 7, 11).Address).columns.AutoFit
                End If
                
            Next
            
            
            '���������G�N�Z�����N���b�Y����B
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '�ċA����������
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracte(subFolder)
    Next
    
    ExcleExtracte = True

    
End Function

'***********************************************************************************************************************
' �@�\   : �e�[�u����`���W��@�\
' �T�v   : �w��f�B���N�g�z����DM���ʕ����A��G�N�Z���ɏW�񂷂�@�\
' ����   : Folder DM���ʕ��̃f�B���N�g���AString �����Ώۂ̃f�[�^�\�[�X�t�@�C���̃p�X
' �߂�l : True : ��������   False : �������s
'***********************************************************************************************************************
Public Function ExcleExtracte_sheetCopy(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    
    
    tableListSheetName = "�ڎ�"
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                
                    dirWorksheet.Copy After:=OPERATION_WORKBOOK.Worksheets(OPERATION_WORKBOOK.Worksheets.Count)
                    
                    tableDefineSheetName = OPERATION_WORKBOOK.Worksheets(OPERATION_WORKBOOK.Worksheets.Count).name
                    
                    fileCounter = OPERATION_WORKBOOK.Sheets(tableListSheetName).UsedRange.Rows.Count

                    '�ڎ�
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = tableDefineSheetName
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirFolderFile.name
                    
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Hyperlinks.Add Anchor:=OPERATION_WORKBOOK.Sheets(tableListSheetName).Range("B" & (fileCounter + 1)), Address:="", SubAddress:="'" & tableDefineSheetName & "'" & "!A1", TextToDisplay:=tableDefineSheetName
                    
                    fileCounter = fileCounter + 1
                
            Next
            
            '�R�s�[�������e��ۑ�����
            'OPERATION_WORKBOOK.Save
            
            '���������G�N�Z�����N���b�Y����B
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '�ċA����������
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracte_sheetCopy(subFolder)
    Next
    
    ExcleExtracte_sheetCopy = True

    
End Function

'***********************************************************************************************************************
' �@�\   : �e�[�u����`���W��@�\
' �T�v   : �w��f�B���N�g�z����DM���ʕ����A��G�N�Z���ɏW�񂷂�@�\
' ����   : Folder DM���ʕ��̃f�B���N�g���AString �����Ώۂ̃f�[�^�\�[�X�t�@�C���̃p�X
' �߂�l : True : ��������   False : �������s
'***********************************************************************************************************************
Public Function ExcleExtracteForDic(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    fileCounter = 1
    
    Dim columnCounter As Long
    columnCounter = 1
    
    tableListSheetName = "�ڎ�"
    
    '�e�[�u���V�[�g�쐬
    tableDefineSheetName = "�����p"
    Call createNewSheet(tableDefineSheetName, OPERATION_WORKBOOK)
    
    
    
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 1) = "����"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 2) = "���ږ���"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 3) = "�K�w"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 4) = "������"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 5) = "���"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 6) = "�o�C�g��"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 7) = "����"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 8) = "����"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 9) = "�J�n�ʒu"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 10) = "�I���ʒu"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 11) = "����"
    
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                If dirWorksheet.name <> "���ڐ���" And Not (dirWorksheet.name Like "*�L����*") Then
                      
                    '�ڎ�
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = dirWorksheet.Cells(3, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirWorksheet.Cells(4, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 4).value = dirWorksheet.Cells(2, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 5).value = dirFolderFile.name
                    
                    '�e�e�[�u���V�[�g
                    i = 8
                    Do While dirWorksheet.Cells(i, 2) <> "" Or dirWorksheet.Cells(i, 14) <> "" Or dirWorksheet.Cells(i, 15) <> "" Or dirWorksheet.Cells(i, 29) <> ""
                       
                        
                        ���� = dirWorksheet.Range("A" & i).value
                        ���ږ��� = dirWorksheet.Range("B" & i).value _
                                 & dirWorksheet.Range("C" & i).value _
                                 & dirWorksheet.Range("D" & i).value _
                                 & dirWorksheet.Range("E" & i).value _
                                 & dirWorksheet.Range("F" & i).value _
                                 & dirWorksheet.Range("G" & i).value _
                                 & dirWorksheet.Range("H" & i).value _
                                 & dirWorksheet.Range("I" & i).value _
                                 & dirWorksheet.Range("J" & i).value _
                                 & dirWorksheet.Range("K" & i).value _
                                 & dirWorksheet.Range("L" & i).value _
                                 & dirWorksheet.Range("M" & i).value
                                 
                        �K�w = dirWorksheet.Range("N" & i).value
                        ������ = dirWorksheet.Range("O" & i).value
                        ��� = dirWorksheet.Range("Z" & i).value
                        �o�C�g�� = dirWorksheet.Range("AC" & i).value & dirWorksheet.Range("AD" & i).value
                        If ��� = "P" Then
                            ���� = dirWorksheet.Range("AE" & i).value & "." & dirWorksheet.Range("AF" & i).value & dirWorksheet.Range("AG" & i).value
                        End If
                        ���� = dirWorksheet.Range("AI" & i).value
                        �J�n�ʒu = dirWorksheet.Range("AJ" & i).value
                        �I���ʒu = dirWorksheet.Range("AK" & i).value
                        ���� = dirWorksheet.Range("AM" & i).value
        
                        
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 1) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 2) = ���ږ���
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 3) = �K�w
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 4) = ������
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 5) = ���
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 6) = �o�C�g��
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 7) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 8) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 9) = �J�n�ʒu
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 10) = �I���ʒu
                        'OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 11) = ����
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 11) = dirFolderFile.name
                        
                        
                        i = i + 1
                        columnCounter = columnCounter + 1
                    Loop
                    
                    fileCounter = fileCounter + 1
                    
                    

                End If
                
            Next
            
            
            '���������G�N�Z�����N���b�Y����B
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '�ċA����������
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracteForDic(subFolder)
    Next
    
    ExcleExtracteForDic = True

    
End Function

