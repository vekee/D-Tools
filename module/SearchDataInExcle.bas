Attribute VB_Name = "SearchDataInExcle"
'***********************************************************************************************************************
' �@�\   : �G�N�Z������f�[�^�𒊏o����@�\
' �T�v   : ����Z���ʒu�̓��e���f�B���N�g���z���̂��ׂăG�N�Z�����璊�o����
' ����   : Folder ���o�Ώۃf�B���N�g���AString ���o�������̓��e
' �߂�l : ��
'***********************************************************************************************************************
Public Function SearchExcleByContentFromDir(dirFolder As Folder, searchByContent As String)
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
        
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        If dirFolderFile.name Like "*.xls" Or dirFolderFile.name Like "*.xlsx" Then
            On Error GoTo nextFile
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True, PASSWORD:="")
            If dirWorkbook.HasPassword Then
                GoTo nextFile
            End If
            '�����Ώۃt�@�C���̑S�V�[�g�Ɍ�������
            For Each dirWorksheet In dirWorkbook.Worksheets
                Dim contentStr As Variant
                For Each contentStr In Split(searchByContent, ",")
                    Dim findRange As Range
                    Dim firstRange As Range
                    Set findRange = dirWorksheet.UsedRange.Find(contentStr)
                    If Not findRange Is Nothing Then
                        '���������P�ڃZ�����������ʃV�[�g�ɏo�͂���
                        Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                        Set firstRange = findRange
                        Do
                            Set findRange = dirWorksheet.UsedRange.FindNext(findRange)
                            If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                                Exit Do
                            End If
                            '���������Z�����������ʃV�[�g�ɏo�͂���
                            Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                        Loop While firstRange.Address <> findRange.Address
                        
                    End If
                    
                Next

            Next
            
            '���������G�N�Z�����N���b�Y����B
            dirWorkbook.Close (False)

        End If
        
nextFile:
    If Err.Number <> 0 Then
        Resume Next
    End If

    Next
    
    For Each subFolder In dirFolder.SubFolders
        Call SearchExcleByContentFromDir(subFolder, searchByContent)
    Next
    
End Function

'***********************************************************************************************************************
' �@�\   : �G�N�Z������f�[�^�𒊏o����@�\
' �T�v   : ����Z���ʒu�̓��e���f�B���N�g���z���̂��ׂăG�N�Z�����璊�o����
' ����   : Folder ���o�Ώۃf�B���N�g���AString ���o�������̃Z���ʒu
' �߂�l : ��
'***********************************************************************************************************************
Public Function SearchExcleByAddressFromDir(dirFolder As Folder, searchByAddress As String)
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
      
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        If dirFolderFile.name Like "*.xl*" Then
            On Error GoTo nextFile
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
            If dirWorkbook.HasPassword Then
                dirWorkbook.Close (False)
                GoTo nextFile
            End If
            '�����Ώۃt�@�C���̑S�V�[�g�Ɍ�������
            For Each dirWorksheet In dirWorkbook.Worksheets
                Dim addressStr As Variant
                For Each addressStr In Split(searchByAddress, ",")
                    Dim findRange As Range
                    Dim firstRange As Range
                    
                    Set findRange = dirWorksheet.Range(addressStr)
                    
                    '���������Z�����������ʃV�[�g�ɏo�͂���
                    Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                                    
                Next
                
            Next

            '���������G�N�Z�����N���b�Y����B
            dirWorkbook.Close (False)
nextFile:
        End If
        
    Next
    
    For Each subFolder In dirFolder.SubFolders
        Call SearchExcleByAddressFromDir(subFolder, searchByAddress)
    Next
End Function
'***********************************************************************************************************************
' �@�\   : �G�N�Z������f�[�^�𒊏o����@�\
' �T�v   : ���o�ΏۃG�N�Z�����璊�o���������e���G�N�Z���Ɋi�[����
' ����   : String ���o�ΏۃG�N�Z���̃f�B���N�g���AString ���o�ΏۃG�N�Z���̃V�[�g���AString ���o�Ώۓ��e�̃Z���ʒu�AString ���o�Ώۓ��e
' �߂�l : ��
'***********************************************************************************************************************
Private Function WriteSearchResultToSheet(filePath As String, sheetName As String, cellAddress As String, cellValue As String)
    Dim usedRowNo As Integer
    Dim writeToRowNo As Integer
    
    usedRowNo = OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Rows.Count
    writeToRowNo = OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Rows(usedRowNo).Row + 1
        
    '�w�b�_���쐬����B
    If usedRowNo = 1 And writeToRowNo = 2 Then
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 1) = "�t�@�C��"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 2) = "�V�[�g��"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 3) = "�ʒu"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 4) = "���e"
    End If
        
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 1) = filePath
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 2) = sheetName
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 3) = cellAddress
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 4) = cellValue
    
    '�ۑ�
    'OPERATION_WORKBOOK.Save
    
End Function
