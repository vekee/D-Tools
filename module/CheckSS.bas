Attribute VB_Name = "CheckSS"
Private Const HEAD_TITLE = "�݌v���`�F�b�N" '�R���e�L�X�g���j���[�̃^�C�g��
Private Const MENU_AUTO = "�݌v���`�F�b�N_����" '�T�u���j���[�̃^�C�g���i�J���}��؂�j
Private Const MENU_MANUAL = "�݌v���`�F�b�N_�蓮�I��" '�T�u���j���[�̃^�C�g���i�J���}��؂�j
Private Const MENU_ACT = "�݌v���`�F�b�N" '�T�u���j���[��Sub���i�J���}��؂�j

Private Const CHECK_INPUT_SHEET_NAME = "���s����_�Z�N�V�����\��"
Private Const CHECK_OBJ_SHEET_NAME = "�������e"

Private Const CHECK_RESULT_SHEET_NAME = "�`�F�b�N����"

'***********************************************************************************************************************
' �@�\   : �݌v���`�F�b�N
' �T�v   : �����������A�ڍא݌v���ɏ������R�ꂪ�Ȃ������`�F�b�N���āA�`�F�b�N���ʂ��V�[�g�ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Function �݌v���`�F�b�N_�蓮�I��() As String

    Dim OpenFileName As String
    
    ChDir "\\cs.dir.co.jp\document\���ʃG���A\HN_Project\PJ_BSNS-20121227-00005_M��G�V��i�p�[�g�i���L�j\30.ZETTA���A���^�C�����E�N���E�h��\99.WORK\90.JS\80_����\18_�N���E�h���istep2�j\03_�����֘A\02_�݌v"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        '�t�@�C���̕����I����s�\�ɂ���
        .AllowMultiSelect = False
        '�t�@�C���t�B���^�̃N���A
        .Filters.Clear
        '�t�@�C���t�B���^�̒ǉ�
        .Filters.Add "�G�N�Z���u�b�N", "*.xls*"
        '�����\���t�H���_�̐ݒ�
        .InitialFileName = "\\cs.dir.co.jp\document\���ʃG���A\HN_Project\PJ_BSNS-20121227-00005_M��G�V��i�p�[�g�i���L�j\30.ZETTA���A���^�C�����E�N���E�h��\99.WORK\90.JS\80_����\18_�N���E�h���istep2�j\03_�����֘A\02_�݌v\"

        If .Show = -1 Then  '�t�@�C���_�C�A���O�\��
            ' [ OK ] �{�^���������ꂽ�ꍇ
            OpenFileName = .SelectedItems(1)
        Else
           End
        End If
    End With

    Call �݌v���`�F�b�N(OpenFileName)

End Function
'***********************************************************************************************************************
' �@�\   : �݌v���`�F�b�N
' �T�v   : �����������A�ڍא݌v���ɏ������R�ꂪ�Ȃ������`�F�b�N���āA�`�F�b�N���ʂ��V�[�g�ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Function �݌v���`�F�b�N(resultWorkBookPath As String)
    
    On Error GoTo errorHandler
    
    '�`�F�b�N���ʃV�[�g�쐬
    '���s���ʂ̏o�͐�
    Dim resultWorkBook As Workbook
    If resultWorkBookPath <> "" Then
        'Workbooks.Open resultWorkBookPath
        'Set resultWorkBook = Application.ActiveWorkbook
        
        Set resultWorkBook = Workbooks.Open(resultWorkBookPath, UpdateLinks:=False, ReadOnly:=True, PASSWORD:="")
        If resultWorkBook.HasPassword Then
            MsgBox "�p�X���[�h���������Ă��������x�����Ă��������I"
        End If
        
    Else
        Set resultWorkBook = Application.ActiveWorkbook
    End If
    
    '���ʏo�͗p�V�[�g���쐬����
    Call createNewSheet(CHECK_RESULT_SHEET_NAME, resultWorkBook)
    
    '�`�F�b�N�Ώۂ̃Z�N�V�����\�����`�F�b�N���ʃV�[�g�ɃR�s�[
    resultWorkBook.Worksheets(CHECK_INPUT_SHEET_NAME).Range("A1").EntireColumn.Copy
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Range("A1").EntireColumn.PasteSpecial (xlPasteAll)
    
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 2).value = "���݌�"
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 3).value = "�`�F�b�N����"
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 4).value = "���l"
    
    
    '�`�F�b�N�Ώۂ��W�񂷂�
    Dim checkMaxRowNo As Long
    Dim sectionNo As String
    
    'checkMaxRowNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows.Count
    checkMaxRowNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows(resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows.Count).Row
    
    For i = 7 To checkMaxRowNo
    
        sectionNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 1).value
        sectionNo = SectionNoFilter(sectionNo)
        
        resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 4).value = sectionNo
        
        If sectionNo <> "" Then
        
            Dim findRange As Range
            Dim firstRange As Range
            Dim findCount As Long
            findCount = 0
            Set findRange = resultWorkBook.Worksheets(CHECK_OBJ_SHEET_NAME).UsedRange.Find(sectionNo)
            
            If Not findRange Is Nothing Then
                '���������P�ڃZ�����������ʃV�[�g�ɏo�͂���
                findCount = findCount + 1
                Set firstRange = findRange
                Do
                    Set findRange = resultWorkBook.Worksheets(CHECK_OBJ_SHEET_NAME).UsedRange.FindNext(findRange)
                    If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                        Exit Do
                    End If
                    '���������Z�����������ʃV�[�g�ɏo�͂���
                    findCount = findCount + 1
                Loop While firstRange.Address <> findRange.Address
                
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 2).value = findCount
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).value = "����"
            Else
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 2).value = findCount
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).value = "�s����"
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).Font.colorIndex = 3
            End If
        Else
            resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 4).value = "�`�F�b�N�ΏۊO"
        End If
        
    Next
    
   
    '�`�F�b�N���ʃV�[�g���`
    '�r����t����
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range(Cells(6, 1).Address & ":" & Cells(i - 1, 4).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range(Cells(6, 1).Address & ":" & Cells(i - 1, 4).Address).columns.AutoFit
    '�I���Z��
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range("A1").Select
    
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("�݌v���`�F�b�N")
    End If
    
End Function

'***********************************************************************************************************************
' �@�\   : sectionNo�����o��
' �T�v   : sectionNo�����o��
' ����   : sectionNo
' �߂�l : sectionNo (�`�F�b�N�ΏۊO�̏ꍇ�A��)
'***********************************************************************************************************************
Private Function SectionNoFilter(sectionNo As String) As String
    sectionNo = Replace(sectionNo, " ", "")
    sectionNo = Replace(sectionNo, "�@", "")
    sectionNo = Replace(sectionNo, "*", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, ".", "")
    sectionNo = Replace(sectionNo, "@", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, "��", "")
    sectionNo = Replace(sectionNo, "��", "")
    
    If sectionNo Like "<*>" Then
        sectionNo = ""
    End If
    
    SectionNoFilter = sectionNo
End Function
'***********************************************************************************************************************
' �@�\   : ���j���[�o�[�ǉ�
' �T�v   : �E�N���b�N�̃��j���[�o�[�ɒǉ�����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Sub AddMenu()
    On Error Resume Next

    '�R���g���[���ɐ݌v���`�F�b�N��ǉ�����
    If Not IsControl(HEAD_TITLE) Then
        Set contextmenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
        With contextmenu
            .Caption = HEAD_TITLE
            .BeginGroup = False
        End With
    End If

    '�T�u�R���g���[���ɐ݌v���`�F�b�N��ǉ�����
    Set contextmenu = Application.CommandBars("Cell").Controls(HEAD_TITLE)
    With contextmenu
        If Not IsSubControl(MENU_AUTO) Then
            With .Controls.Add(Temporary:=True)
                .Caption = MENU_AUTO
                .OnAction = MENU_ACT
                .Enabled = False
            End With
        End If
    End With
    '�T�u�R���g���[���ɐ݌v���`�F�b�N��ǉ�����
    With contextmenu
        If Not IsSubControl(MENU_MANUAL) Then
            With .Controls.Add(Temporary:=True)
                .Caption = MENU_MANUAL
                .OnAction = MENU_ACT & "_�蓮�I��"
            End With
        End If
    End With
End Sub
'***********************************************************************************************************************
' �@�\   : �������ʋ@�\
' �T�v   : �R���g���[�������݂��邩�ǂ����`�F�b�N����
' ����   : String �R���g���[��
' �߂�l : Boolean TRUE�F���݁^FALSE�F�s����
'***********************************************************************************************************************
Private Function IsControl(name As String) As Boolean
    Dim found As Boolean

    For Each C In Application.CommandBars("Cell").Controls
        If C.Caption = name Then
            found = True
            Exit For
        End If
    Next C
    IsControl = found
End Function
'***********************************************************************************************************************
' �@�\   : �������ʋ@�\
' �T�v   : �T�u�R���g���[�������݂��邩�ǂ����`�F�b�N����
' ����   : String �T�u�R���g���[��
' �߂�l : Boolean TRUE�F���݁^FALSE�F�s����
'***********************************************************************************************************************
Private Function IsSubControl(name As String) As Boolean
    On Error GoTo ex
    Dim found As Boolean
    found = False
    For Each C In Application.CommandBars("Cell").Controls(HEAD_TITLE).Controls
        If C.Caption = name Then
            found = True
            Exit For
        End If
    Next C

ex:
    IsSubControl = found
End Function
