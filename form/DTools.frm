VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DTools 
   Caption         =   "D-Tools"
   ClientHeight    =   11130
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   17410
   OleObjectBlob   =   "DTools.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "DTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************************
' �@�\   : D-Tools�������@�\
' �T�v   : �A�h�C�����瑀�엚�������擾���āA��ʂ֐ݒ肷��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub UserForm_Initialize()

    Set addInWorkBook = Application.ThisWorkbook
    Set OPERATION_WORKBOOK = Application.ActiveWorkbook
    
    '�ڑ����V�[�g�쐬
    Set addInConnInfoWS = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME)
    
    '���엚���V�[�g������e��ݒ肷��
    '���sSQL
    DTools.sqlTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(1, 3).value
    '�e�[�u��������
    DTools.GetTableLayoutTableNameTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value
    'InsertSql
    DTools.InsertSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value
    'UpdateSql
    DTools.UpdateSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value
    'DeleteSql
    DTools.DeleteSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value

    '�f�B���N�g��
    DTools.DirTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(9, 3).value
    '���o����
    DTools.GetContentTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(10, 3).value
    '����̓��e
    DTools.GetByContentOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(11, 3).value
    '����̈ʒu
    DTools.GetByAddressOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(12, 3).value
    'D-Tools��`���|�W�g��
    DTools.DMRepositoryTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(13, 3).value
    '�e�[�u����`��
    DTools.TableCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value
    'DM��`���ʕ��i�[�ꏊ
    DTools.LatestDMDefineFileTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value
    'DB��`
    DTools.TableInfoInDaBaseOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value
    'D-Tools��`���|�W�g��
    DTools.TableInfoInDMRepositoryOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value
    

    
    '�J�n�Z��
    DTools.SetColorInCellStartTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(24, 3).value
    '�I���Z��
    DTools.SetColorInCellEndTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(25, 3).value
    '�����J�n
    DTools.SetColorInCellCharStartTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(26, 3).value
    '�F�t��������
    DTools.SetColorInCellCharLengTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(27, 3).value
    '��
    DTools.SetColorInCellRedColorOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(28, 3).value
    '��
    DTools.SetColorInCellBlueColorOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(29, 3).value
    

    
    'DB�ڑ�comboBox�쐬
    DTools.DBConnInfoComboBox.Clear
    CONN_INFO_MAX_ROW = addInConnInfoWS.UsedRange.Rows.Count
    For i = 1 To CONN_INFO_MAX_ROW
        DTools.DBConnInfoComboBox.AddItem (addInConnInfoWS.Cells(i, 1))
        'TextBox�̏����l��ݒ肷��
        If addInConnInfoWS.Cells(i, 3) = SELECT_ON Then
            DTools.DBConnInfoComboBox.ListIndex = i - 1
            DTools.DBConnInfoTextBox.Text = addInConnInfoWS.Cells(i, 2)
        End If
    Next
    
    '�t�H�[�������t���[�V�����邽��
    'Call saveConnInfoButton_Click
    
    '�ڑ�����������
    DB_CONN_INFO_STR = DTools.DBConnInfoTextBox.Text
    '�ϐ�������
    DATA_SOURCE_DIR = DTools.DMRepositoryTextBox.Text
End Sub
'***********************************************************************************************************************
' �@�\   : D-Tools��ʂ����@�\
' �T�v   : �u����v�{�^�����������鎞�AD-Tools��ʂ����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub closeFormButton_Click()
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : �ڑ����ۑ��@�\
' �T�v   : �ڑ������͂̃h���b�v�_�E�����X�g�̕ύX�𔭐��������A�ڑ������X�V����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub DBConnInfoComboBox_Change()
    '�ڑ�����ύX����
    CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
    For i = 1 To CONN_INFO_MAX_ROW
        Dim itemValue As String
        itemValue = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1)
        If DTools.DBConnInfoComboBox.Text = itemValue Then
            DTools.DBConnInfoTextBox.Text = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
            addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
            addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON
            
            '�S�̗p�ϐ��Đݒ�
            DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
            
        End If
    Next
End Sub
'***********************************************************************************************************************
' �@�\   : �ڑ����ۑ��@�\
' �T�v   : �ڑ������͂̃e�L�X�g�{�b�N�X�̕ύX�𔭐��������A�ڑ����ۑ��������I�ɍs��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub DBConnInfoTextBox_Change()
    Call saveConnInfoButton_Click
End Sub

'***********************************************************************************************************************
' �@�\   : �ڑ����ۑ��@�\
' �T�v   : �u�ڑ����ۑ��v�{�^�����������鎞�A���͂����ڑ������A�h�C���Ɋi�[����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub saveConnInfoButton_Click()
    Dim DBConnInfoTextBox As String
    Dim DBConnInfoComboBox As ComboBox
    Dim comboBoxSelectedText As String
    Dim existFlag As Boolean
    Dim itemValue As String
    existFlag = False
    DBConnInfoTextBox = DTools.DBConnInfoTextBox.Text
    
    Set DBConnInfoComboBox = DTools.DBConnInfoComboBox
    comboBoxSelectedText = DBConnInfoComboBox.Text
    
    If comboBoxSelectedText <> "" Or DBConnInfoTextBox <> "" Then
        
        If comboBoxSelectedText = "" Then
            comboBoxSelectedText = "(��)"
        End If
    
    
        '�f�[�^�i�[
        CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
        
        '�����̐ڑ���ύX����ꍇ
        For i = 1 To CONN_INFO_MAX_ROW
            If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1) = comboBoxSelectedText Then
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2) = DBConnInfoTextBox
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON
                '�S�̗p�ϐ��Đݒ�
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                '��������t���O
                existFlag = True
            End If
        Next
        
        '�V�K�̐ڑ����쐬����ꍇ
        If existFlag = False Then
            If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 1) <> "" Then
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 2) = DBConnInfoTextBox
                '�S�̗p�ϐ��Đݒ�
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 3) = SELECT_ON
            Else
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 2) = DBConnInfoTextBox
                '�S�̗p�ϐ��Đݒ�
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 3) = SELECT_ON
            End If
        End If

        '�\�[�g
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").Sort Key1:=addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:A"), order1:=xlAscending, Header:=xlNo, MatchCase:=False, SortMethod:=xlPinYin
       
        '�d���ڑ��̃f�[�^���폜����
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").RemoveDuplicates columns:=Array(1, 2), Header:=xlNo
        
        'comboBox�č쐬
        DBConnInfoComboBox.Clear
        CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
        For i = 1 To CONN_INFO_MAX_ROW
            itemValue = ""
            itemValue = addInConnInfoWS.Cells(i, 1)
            DBConnInfoComboBox.AddItem (itemValue)
            '�ۑ��l��\���ɐݒ肷��
            If comboBoxSelectedText = itemValue Then
                DBConnInfoComboBox.ListIndex = i - 1
            End If
        Next
        
         Set DBConnInfoComboBox = DTools.DBConnInfoComboBox
        
    End If
    
    If comboBoxSelectedText = "" And DBConnInfoTextBox = "" Then
        '�ڑ������폜����
        'comboBox�č쐬
        DBConnInfoComboBox.Clear
        CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
        For i = 1 To CONN_INFO_MAX_ROW
            itemValue = ""
            itemValue = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1)
            If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON Then
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1) = ""
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2) = ""

            Else
                DBConnInfoComboBox.AddItem (itemValue)
            End If
        Next
        
        '�ڑ����X�g�̈�ڂ�\������
        DBConnInfoComboBox.ListIndex = 0
        
        '�\�[�g
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").Sort Key1:=addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A1"), order1:=xlAscending, Header:=xlNo, MatchCase:=False, SortMethod:=xlPinYin
        '�d���ڑ��̃f�[�^���폜����
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").RemoveDuplicates columns:=Array(1, 2), Header:=xlNo
        
        '�ύX�̐ڑ�����ۑ�����
        'addInWorkBook.Save
        
    End If
          
End Sub

'***********************************************************************************************************************
' �@�\   : SQL���s�@�\
' �T�v   : �u���s�v�{�^�����������鎞�A���͂���SQL�����s���āA���s���ʂ��G�N�Z���ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub runSQLButton_Click()
    Dim sqls As String
    sqls = DTools.sqlTextBox.Text
    
    '���͏���ۑ�����
    '���sSQL
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(1, 3).value = sqls
    
    'SQL�̋󔒃`�F�b�N
    If Replace(Replace(sqls, " ", ""), "�@", "") = "" Then
        Exit Sub
    End If
        
    '���s���폜����
    sqls = Replace(sqls, ";" & vbCrLf, ";")
    
    '���s���ʂ̏o�͐�
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '���ʏo�͗p�V�[�g���쐬����
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connDB()
    
    On Error GoTo errorHandler
    
    Dim rs As New ADODB.recordset
    rowIndex = 0
    Dim sql As Variant
    For Each sql In Split(sqls, ";")
        'SQL�̋󔒃`�F�b�N
        If Replace(Replace(sql, " ", ""), "�@", "") = "" Then
            GoTo Continue
        End If
    
        colIndex = 1
        rowIndex = rowIndex + 2
        '���ʏW������������
        
        'sql�̕ҏW����B
        
        Dim recordsAffected As Long
        Set rs = New ADODB.recordset
        Set rs = ADOConnection.Execute(sql, recordsAffected)
        
        If rs.State = 0 Then
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 1).value = sql
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 2).value = recordsAffected & "�����R�[�h���e�����܂����B"
            '�r����t����
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, 2).Address).Borders.LineStyle = xlContinuous
            GoTo Continue
        End If
        
        Dim resultFields As ADODB.Fields
        Dim resultField As ADODB.field
        Set resultFields = rs.Fields
        
        '�J���������o�͂���
        For Each resultField In resultFields
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, colIndex).value = resultField.name
            colIndex = colIndex + 1
        Next
        '�r����t����
        resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, colIndex - 1).Address).Borders.LineStyle = xlContinuous
    
        '�f�[�^���o�͂���
        Do Until rs.EOF
            rowIndex = rowIndex + 1
            colIndex = 1

            For Each resultField In resultFields
                resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, colIndex).value = resultField.value
                colIndex = colIndex + 1
            Next
            
            '�r����t����
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, colIndex - 1).Address).Borders.LineStyle = xlContinuous
            
            rs.MoveNext
        Loop

Continue:
    Next
    
    If rs.State = 1 Then
        rs.Close
    End If
    ADOConnection.Close
    
    '��̕���������
    resultWorkBook.Sheets(RESULT_SHEET_NAME).UsedRange.columns.AutoFit
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("runSQLButton_Click")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If

End Sub

'***********************************************************************************************************************
' �@�\   : ���s�v��擾�@�\
' �T�v   : �u���s�v��v�{�^�����������鎞�A���͂���SQL�̎��s���ʂ��擾���āA�G�N�Z���ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub SqlExecutePlanCommandButton_Click()
    
    '���͏���ۑ�����
    '���sSQL
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(1, 3).value = DTools.sqlTextBox.Text
    
    Dim sql As String
    Dim sqlId As String
    sql = DTools.sqlTextBox.Text
    sql = Replace(sql, ";", " ")
    
    
    Dim sqlKeyWord As String
    sqlKeyWord = Format(Now, "yyyymmddhhmmss") & Environ("COMPUTERNAME")
    
    
    Dim explainSql As String
    explainSql = "EXPLAIN PLAN FOR " & sql & " /*" & sqlKeyWord & "*/"
    
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Set ADOConnection = Utils.connDB()
    
    On Error GoTo errorHandler
    
    '�@���s�v�����͂���
    ADOConnection.Execute explainSql
    
    '�ASQLID���擾����
    'Dim getSqlId As String
    'getSqlId = "SELECT sql_id FROM v$sql WHERE sql_text LIKE '%" & sqlKeyWord & "%' AND ROWNUM <=1"
    'Set ADORecordset = New ADODB.recordset
    'Set ADORecordset = ADOConnection.Execute(getSqlId)
    'Do Until ADORecordset.EOF
    '    For Each field In ADORecordset.Fields
    '        sqlId = field.value
    '    Next
    '    ADORecordset.MoveNext
    'Loop
    
    '���s�v����擾����
    Dim getExplainSql As String
    'getExplainSql = "SELECT * FROM TABLE(DBMS_XPLAN.DISPLAY_CURSOR('" & sqlId & "',0))"
    getExplainSql = "SELECT * FROM TABLE(DBMS_XPLAN.DISPLAY())"
    
    Set ADORecordset = New ADODB.recordset
    Set ADORecordset = ADOConnection.Execute(getExplainSql)
    
    Dim result As String
    Do Until ADORecordset.EOF
        For Each field In ADORecordset.Fields
            result = result & field.value
        Next
        result = result & vbCrLf
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("SqlExecutePlanCommandButton_Click")
    Else
        If result <> "" Then
            DTools.sqlTextBox.Text = sql & vbCrLf _
                                        & "/**************************************************************************************" _
                                        & vbCrLf & vbCrLf _
                                        & result & vbCrLf _
                                        & "**************************************************************************************/ "
        End If
        
    End If
    
    '���[�U�[�̐ݒ�̏���ۑ�����B
    'addInWorkBook.Save
    
End Sub

'***********************************************************************************************************************
' �@�\   : �J�����Q�Ƌ@�\
' �T�v   : �u�J�����Q�Ɓv�{�^�����������鎞�A���͂����e�[�u�����������A�e�[�u�����C�A�E�g�����擾���āA�G�N�Z���ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub getTableLayoutButton_Click()
    
    '���͏���ۑ�����
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value = DTools.GetTableLayoutTableNameTextBox.Text
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value = DTools.TableInfoInDaBaseOptionButton.value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value = DTools.TableInfoInDMRepositoryOptionButton.value
    
    '���̓`�F�b�N
    If DTools.GetTableLayoutTableNameTextBox.Text = "" Then
        MsgBox "�e�[�u��������͂��Ă��������B"
        Exit Sub
    End If

    'On Error GoTo errorHandler
    
    '���ʏo�͗p�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    Dim tableName As Variant
    Dim rowIndex As Integer
    rowIndex = 1
    For Each tableName In Split(DTools.GetTableLayoutTableNameTextBox.Text, ",")
        Dim tableNameInfoCollection As New Collection
        
        '�e�[�u���������擾����B
        If DTools.TableInfoInDaBaseOptionButton.value = True Then
            Set tableNameInfoCollection = GetTableNameFromDB(CStr(tableName))
        Else
            Set tableNameInfoCollection = GetTableNameFromDMRepository(CStr(tableName))
       
        End If
        
        If tableNameInfoCollection.Count > 0 Then

            Dim columnsInfoCollection As New Collection
            Dim columnInfo() As String
            tableNameIndex = 1
            Do While tableNameIndex <= tableNameInfoCollection.Count
                Dim tableNameInfo() As String
                ReDim tableNameInfo(1)
                tableNameInfo = tableNameInfoCollection(tableNameIndex)
                
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 1).value = tableNameInfo(1)
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 2).value = tableNameInfo(0)
                '�r����t����
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address, Cells(rowIndex, 2).Address).Borders.LineStyle = xlContinuous
                
                '�e�[�u���J���������擾����
                If DTools.TableInfoInDaBaseOptionButton.value = True Then
                    Set columnsInfoCollection = GetTableColumnsNameFromDB(tableNameInfo(0))
                Else
                    Set columnsInfoCollection = GetTableColumnsNameFromDMRepository(tableNameInfo(0))
                End If
                
                colIndex = 1
                Do While colIndex <= columnsInfoCollection.Count
                    Dim columnName() As String
                    ReDim columnName(5)
                    columnName = columnsInfoCollection(colIndex)
                    
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 1, colIndex).value = columnName(0)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 2, colIndex).value = columnName(1)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 3, colIndex).value = columnName(2)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 4, colIndex).value = columnName(3)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 5, colIndex).value = columnName(4)
                    If columnName(5) <> "" Then
                        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 1, colIndex).Interior.Color = RGB(279, 117, 14)
                    End If
                    colIndex = colIndex + 1
                Loop
                
                '�r����t����
                 OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex + 1, 1).Address & ":" & Cells(rowIndex + 5, colIndex - 1).Address).Borders.LineStyle = xlContinuous
                '��̕���������
                 OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex + 5, colIndex - 1).Address).columns.AutoFit
        
                rowIndex = rowIndex + 8
                
                tableNameIndex = tableNameIndex + 1
            Loop
            
        End If
        
    Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("getTableLayoutButton_Click")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If
    
End Sub
'***********************************************************************************************************************
' �@�\   : �����f�[�^�쐬�@�\
' �T�v   : �u�����f�[�^�쐬�v�{�^�����������鎞�A��Ɨp�V�[�g���쐬���āA�����f�[�^�쐬���s��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateTestDataCommandButton_Click()
    '���͏���ۑ�����
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value = DTools.GetTableLayoutTableNameTextBox.Text
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value = DTools.TableInfoInDaBaseOptionButton.value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value = DTools.TableInfoInDMRepositoryOptionButton.value
       
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools��`���|�W�g����ݒ肵�Ă��������B"
        Exit Sub
    End If
    
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "�쐬�Ώۃe�[�u��"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(2, 1) = "���R�[�h��"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(3, 1) = "�l���ʎq" & vbCrLf & "�i�ԍ��擪��3�`4���j"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(4, 1) = "�ԍ��̎}��" & vbCrLf & "�i�ԍ�������3���j"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 1) = "�w��_���J������" & vbCrLf & "(�����\���L�ډ�)"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 2) = "�w��l"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 5) = "���p�Җ�"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 6) = "�l���ʎq" & vbCrLf & "�i�ԍ��擪��3�`4���j"

    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(2, 2) = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(19, 3).value
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(3, 2) = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(20, 3).value
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(4, 2) = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(21, 3).value
    
    Dim columnNameJP As Variant
    Dim columnValue As Variant
    Dim dataRowStartIndex As Integer
    dataRowStartIndex = 6
    For Each columnNameJP In Split(addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(22, 3).value, ",")
        resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(dataRowStartIndex, 1) = columnNameJP
        dataRowStartIndex = dataRowStartIndex + 1
    Next
    dataRowStartIndex = 6
    For Each columnValue In Split(addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(23, 3).value, ",")
        resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(dataRowStartIndex, 2) = columnValue
        dataRowStartIndex = dataRowStartIndex + 1
    Next
    
    Dim columnValueKeyInfoCollection As New Collection
    Dim columnValueKeyInfo As Variant
    Set columnValueKeyInfoCollection = GetKeyValue("2")
    dataRowStartIndex = 6
    For Each columnValueKeyInfo In columnValueKeyInfoCollection
        resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(dataRowStartIndex, 5) = columnValueKeyInfo(0)
        resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(dataRowStartIndex, 6) = columnValueKeyInfo(1)
        dataRowStartIndex = dataRowStartIndex + 1
    Next
    
    Dim usedRowCount As Integer
    usedRowCount = OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Rows.Count
    
    If usedRowCount > 20 Then
        '�r����t����
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(usedRowCount, 2).Address).Borders.LineStyle = xlContinuous
    Else
        '�r����t����
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(20, 2).Address).Borders.LineStyle = xlContinuous
    End If
    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).Borders.LineStyle = xlContinuous
    
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A5").Interior.Color = RGB(255, 153, 0)
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B5").Interior.Color = RGB(255, 153, 0)
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).Interior.Color = RGB(128, 128, 128)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 45
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A2:A5").RowHeight = 30

    
    '�����̈ʒu�𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").HorizontalAlignment = xlLeft
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").VerticalAlignment = xlTop
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A4").HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A4").VerticalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(5).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(5).VerticalAlignment = xlCenter
    
    '��̕��𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 55
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("D1").ColumnWidth = 20
    
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(usedRowCount, 1).Address).columns.AutoFit
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).columns.AutoFit

    '�u�쐬����v�{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("D1").Left, _
                                 Range("D1").Top, _
                                 Range("D1").Width, _
                                 Range("D1").Height)
        .OnAction = "CreateTestData"
        .Characters.Text = "�쐬����"
    End With
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' �@�\   : SQL�쐬�@�\
' �T�v   : �uSQL�쐬�v�{�^�����������鎞�A�I�����ꂽSQL�쐬��ʂ��ASQL�쐬���āA�t�@�C���ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateSqlCommandButton_Click()
    
    '���͏���ۑ�����
    'InsertSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value = DTools.InsertSqlCheckBox.value
    'UpdateSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value = DTools.UpdateSqlCheckBox.value
    'DeleteSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value = DTools.DeleteSqlCheckBox.value
    
    '�`�F�b�N
    If ActiveSheet.UsedRange.Rows.Count < 7 Then
        MsgBox ("�e�[�u�����C�A�E�g�i�����s���I DTools�ŏo�͂����e�[�u�����C�A�E�g�𗘗p���Ă��������I")
    End If
    
    
    '�S�̒萔�𗘗p����
    Set PUB_TEMP_VAR_OBJ = CreateObject("ADODB.Stream")
    PUB_TEMP_VAR_OBJ.Type = adTypeText
    PUB_TEMP_VAR_OBJ.Charset = "UTF-8"
    PUB_TEMP_VAR_OBJ.LineSeparator = adCRLF
    PUB_TEMP_VAR_OBJ.Open
    
    Dim columnsCount As Integer
    columnsCount = ActiveSheet.UsedRange.columns.Count
    Dim usedRowsEndIndex As Integer
    usedRowsEndIndex = ActiveSheet.UsedRange.Rows.Count
    
    '�f�[�^����z��Ɋi�[����
    tableNameJP = ActiveSheet.UsedRange.Cells(1, 1).value
    tableNameEN = ActiveSheet.UsedRange.Cells(1, 2).value
    '�X�L�[�}���͔��f
    If UBound(Split(tableNameEN, ".")) = 0 Then
        tableNameEN = "MI_TRAN." & tableNameEN
    End If
    columnNameEnArray = WorksheetFunction.Transpose(ActiveSheet.UsedRange.Rows(3))
    columnInfoSetArray = ActiveSheet.UsedRange.Rows("1:6")

    Dim dataSetArray() As Variant
    ReDim dataSetArray(usedRowsEndIndex - 7)
    For i = 7 To usedRowsEndIndex
        dataSetArray(i - 7) = WorksheetFunction.Transpose(ActiveSheet.UsedRange.Rows(i))
    Next
    
    'UpdateSql�ADeleteSql���쐬���鎞�AWhere�������쐬����
    Dim whereSqlSetArray As Variant
    If DTools.UpdateSqlCheckBox.value = True Or DTools.DeleteSqlCheckBox.value = True Then
        whereSqlSetArray = CreateWhereSql()
    End If
    
     'UpdateSql���쐬����
     If DTools.UpdateSqlCheckBox.value = True Then
        Call CreateUpdateSqlSimple(tableNameJP, tableNameEN, columnNameEnArray, dataSetArray, whereSqlSetArray)
     End If
     
     'DeleteSql���쐬����
     If DTools.DeleteSqlCheckBox.value = True Then
        Call CreateDeleteSqlSimple(tableNameJP, tableNameEN, whereSqlSetArray)
     End If
     
     'InsertSql���쐬����
     If DTools.InsertSqlCheckBox.value = True Then
         Call CreateInsertSqlSimple(tableNameJP, tableNameEN, columnNameEnArray, dataSetArray)
     End If
     
    Dim sqlFile As String
    sqlFile = GetSaveDir & "\" & tableNameJP & "_" & Format(Now, "yyyymmddHHMMSS") & ".sql"
       '�t�@�C����ۑ�����
    PUB_TEMP_VAR_OBJ.SaveToFile (sqlFile), adSaveCreateOverWrite
    '�t�@�C���ƕ���
    PUB_TEMP_VAR_OBJ.Close

    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
    MsgBox "SQL���쐬�������܂����B" & vbCrLf & "�i�[�ꏊ�F" & sqlFile
    
End Sub

'***********************************************************************************************************************
' �@�\   : SQL�쐬�@�\
' �T�v   : �uSQL�쐬�v�{�^�����������鎞�A�I�����ꂽSQL�쐬��ʂ��ASQL�쐬���āA�t�@�C���ɏo�͂���
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateSqlCommandButton_Click_BK()
    Dim usedRowsCount As Integer
    Dim usedRowsEndIndex As Integer
    Dim usedRowsStartIndex As Integer
    
    Dim dataRowsStartIndex As Integer


    '���͏���ۑ�����
    'InsertSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value = DTools.InsertSqlCheckBox.value
    'UpdateSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value = DTools.UpdateSqlCheckBox.value
    'DeleteSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value = DTools.DeleteSqlCheckBox.value
    
    
    If DTools.InsertSqlCheckBox.value = False And DTools.UpdateSqlCheckBox.value = False And DTools.DeleteSqlCheckBox.value = False Then
        MsgBox "�쐬����Sql��ʂ�I�����Ă��������B"
        Exit Sub
    End If
    
    '�S�̒萔�𗘗p����
    Set PUB_TEMP_VAR_OBJ = CreateObject("ADODB.Stream")
    PUB_TEMP_VAR_OBJ.Type = adTypeText
    PUB_TEMP_VAR_OBJ.Charset = "UTF-8"
    PUB_TEMP_VAR_OBJ.LineSeparator = adCRLF
    PUB_TEMP_VAR_OBJ.Open

    
    usedRowsCount = ActiveSheet.UsedRange.Rows.Count
    usedRowsEndIndex = ActiveSheet.UsedRange.Rows(usedRowsCount).Row
    usedRowsStartIndex = usedRowsEndIndex - usedRowsCount + 1
    
    Dim rowCounter As Integer
    rowCounter = usedRowsStartIndex
    
    Dim sqlFileName As String
    Dim tableNameEN As String
    Dim getTableColumnsResult As Variant
    Dim tableColumns() As String
    Dim columnStartIndex As Integer
    Dim columnEndIndex As Integer
    Dim dataCollection As New Collection
    Dim countinue As Boolean
    Dim oneSetFinish As Boolean

    Do While rowCounter <= usedRowsEndIndex
        '�ϐ�������
        countinue = False
        oneSetFinish = False
        
        If tableNameEN = "" Then
            If RowIsAllSpace(rowCounter) = True Then
                '���̍s��T���ɍs��
                countinue = True
            Else
                '�e�[�u�����������擾����B
                tableNameEN = CreateSql_GetTableName(rowCounter)
            End If
            
            '�擾���ʂ̃`�F�b�N
            If countinue <> True And tableNameEN = "" Then
                MsgBox rowCounter & "�s�ڂ̃f�[�^�i�����s���A���m�F���������B"
                Exit Sub
            Else
                '���̍s��T���ɍs��
                countinue = True
            End If
        End If
        
        '�e�[�u���J���������擾����B
        If countinue <> True And tableNameEN <> "" And IsArrayEx(tableColumns) = 0 Then
            If RowIsAllSpace(rowCounter) = True Then
                '���̍s��T���ɍs��
                countinue = True
            Else
                ReDim getTableColumnsResult(2)
                getTableColumnsResult = CreateSql_GetTableColumns(rowCounter)
                
                
                columnStartIndex = getTableColumnsResult(1)
                columnEndIndex = getTableColumnsResult(2)
                
                ReDim tableColumns(columnEndIndex - columnStartIndex + 1)
                tableColumns = getTableColumnsResult(0)
                
            End If
            
            
            '�擾���ʂ̃`�F�b�N
            If countinue <> True And IsArrayEx(tableColumns) = 0 Then
                '���̍s��T���ɍs��
                countinue = True
            End If
            If countinue <> True And IsArrayEx(tableColumns) > 0 Then
                '���̍s��T���ɍs��
                countinue = True
            End If
        
        End If
        
        '�f�[�^���W�߂�
        If countinue <> True And tableNameEN <> "" And IsArrayEx(tableColumns) = 1 Then
            If RowIsAllSpace(rowCounter) = True Then
                '���̍s��T���ɍs��
                countinue = True
            ElseIf checkTableNameExistInRow(rowCounter) = True Then
                '������
                tableNameEN = ""
                Erase getTableColumnsResult
                columnStartIndex = 0
                columnEndIndex = 0
                Erase tableColumns
                Set dataCollection = New Collection
    
                '���̍s��T���ɍs��
                countinue = True
            ElseIf checkStrsExistInRow(rowCounter, Array("CHAR", "VACHAR2", "NUMBER")) <> "0" Then
                '�f�[�^�^�A�T�C�Y�ANULL�ۂ��΂��āA�f�[�^�s��T���Č��ɍs���B
                rowCounter = rowCounter + 2
                '���̍s��T���ɍs��
                countinue = True
            Else
                dataCollection.Add (CreateSql_GetData(rowCounter, columnStartIndex, columnEndIndex))
            End If
            
            '�擾���ʂ̃`�F�b�N
            If countinue <> True And dataCollection.Count = 0 Then
                MsgBox rowCounter & "�s�ڂ̃f�[�^�i�����s���A���m�F���������B"
                Exit Sub
            End If
            
            If countinue <> True And dataCollection.Count > 0 And RowIsAllSpace(rowCounter + 1) = True Then
                '��Z�b�g��ݒ芮��
                oneSetFinish = True
            End If
            
        End If
        
        
        '�o��
        If countinue <> True And oneSetFinish = True Then
            
            '�e�[�u�����Ƃ̖��̂��W�߂���
            sqlFileName = sqlFileName & tableNameEN & "_"
        
            'InsertSql���쐬����
            If DTools.InsertSqlCheckBox.value = True Then
                Call CreateInsertSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            'UpdateSql���쐬����
            If DTools.UpdateSqlCheckBox.value = True Then
                Call CreateUpdateSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            'DeleteSql���쐬����
        
            If DTools.DeleteSqlCheckBox.value = True Then
                Call CreateDeleteSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            
            '������
            tableNameEN = ""
            Erase getTableColumnsResult
            columnStartIndex = 0
            columnEndIndex = 0
            Erase tableColumns
            Set dataCollection = New Collection
    
            '���̍s��T���ɍs��
            countinue = True
           
        End If
            
        rowCounter = rowCounter + 1
        
    Loop


    Dim sqlFile As String
    sqlFile = GetSaveDir & "\" & sqlFileName & Format(Now, "yyyymmdd") & ".sql"
       '�t�@�C����ۑ�����
    PUB_TEMP_VAR_OBJ.SaveToFile (sqlFile), adSaveCreateOverWrite
    '�t�@�C���ƕ���
    PUB_TEMP_VAR_OBJ.Close
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("CreateSqlCommandButton_Click")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
        MsgBox "SQL���쐬�������܂����B" & vbCrLf & "�i�[�ꏊ�F" & sqlFile
    End If
    
End Sub
'***********************************************************************************************************************
' �@�\   : �G�N�Z������f�[�^���o�@�\
' �T�v   : �u���o����v�{�^�����������鎞�A�w��ꏊ�̂��ׂăG�N�Z������w��̓��e�𒊏o����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub GetContentCommandButton_Click()
    Dim objFSO As FileSystemObject
    Dim dirFolder As Folder
    
    Dim dir As String
    dir = DTools.DirTextBox.Text

    '���͏���ۑ�����
    '�f�B���N�g��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(9, 3).value = dir
    '���o����
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(10, 3).value = DTools.GetContentTextBox.Text
    '����̓��e
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(11, 3).value = DTools.GetByContentOptionButton.value
    '����̈ʒu
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(12, 3).value = DTools.GetByAddressOptionButton.value
    
    If dir = "" Then
        MsgBox "�����̃f�B���N�g������͂��Ă��������B"
        Exit Sub
    End If
    
    If DTools.GetContentTextBox.Text = "" Then
        MsgBox "�����̃f�[�^����͂��Ă��������B"
        Exit Sub
    End If
    
    If DTools.GetByContentOptionButton.value = False And DTools.GetByAddressOptionButton.value = False Then
        MsgBox "��������(���e�A�Z���ʒu)���w�肵�Ă��������B"
        Exit Sub
    End If
    
    'On Error GoTo errorHandler
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set dirFolder = objFSO.GetFolder(dir)
    
    '���s���ʂ̏o�͐�
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '���ʏo�͗p�V�[�g���쐬����
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    If DTools.GetByContentOptionButton.value = True Then
        
        Call SearchExcleByContentFromDir(dirFolder, DTools.GetContentTextBox.Text)
        
    ElseIf DTools.GetByAddressOptionButton.value = True Then
    
        Call SearchExcleByAddressFromDir(dirFolder, DTools.GetContentTextBox.Text)
    
    End If
    
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 4).Address).Interior.Color = RGB(255, 153, 0)
    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.columns.AutoFit
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("GetContentCommandButton_Click")
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
    End If
    
End Sub
'***********************************************************************************************************************
' �@�\   : �͐ߎ����O���[�v���@�\
' �T�v   : ���[�N�V�[�g�̓��e�������I�ɃO���[�v������
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
'
Private Sub GroupWorkSheetCommandButton_Click()
    usedRangeCount = ActiveSheet.UsedRange.Rows.Count
    
    Dim groupStartIndex As Integer
    Dim groupEndIndex As Integer
    
    Dim firstCellInRow As Range
    Dim endCellInColumn As Range
    
    Dim i As Integer
    i = 1
    Do While i <= usedRangeCount
    
    On Error GoTo nextLoop
        If Not RowIsAllSpace(i) Then
            If Not (GetNotNullCellInOneRow(i) Is Nothing) Then
                Set firstCellInRow = GetNotNullCellInOneRow(i)
                If IsGroupStartRow(firstCellInRow.value) Then
                    Set endCellInColumn = GetNotNullCellUnderOneCell(firstCellInRow)
                    If endCellInColumn.Row - 1 > firstCellInRow.Row + 1 Then
                        groupStartIndex = firstCellInRow.Row + 1
                        groupEndIndex = endCellInColumn.Row - 1
                        Call Group(groupStartIndex, groupEndIndex)
                    End If

                End If
            End If
        End If
nextLoop:
        i = i + 1
    Loop
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' �@�\   : D-Tools��`���|�W�g���ݒ�@�\
' �T�v   : �u�t�@�C���I���v�{�^�����������鎞�A�A�N�Z�X�t�@�C���̑I����ʂ��N������
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
'
Private Sub FindDMRepositoryCommandButton_Click()
    Dim OpenFileName As String, FileName As String
    OpenFileName = Application.GetOpenFilename("Microsoft Access �f�[�^�x�[�X(*.accdb),*.accdb?")
    If OpenFileName <> "False" Then
        DTools.DMRepositoryTextBox.Text = OpenFileName
    End If
End Sub

'***********************************************************************************************************************
' �@�\   : D-Tools��`���|�W�g���ݒ�@�\
' �T�v   : �u�ۑ�����v�{�^�����������鎞�AD-Tools��`���|�W�g���t�@�C�����A�h�C���Ɋi�[����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
'
Private Sub SetDMRepositoryCommandButton_Click()
    '���͏���ۑ�����
    'D-Tools��`���|�W�g��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(13, 3).value = DTools.DMRepositoryTextBox.Text
    DATA_SOURCE_DIR = DTools.DMRepositoryTextBox.Text
End Sub

'***********************************************************************************************************************
' �@�\   : DM�������@�\
' �T�v   : �u��������v�{�^�����������鎞�A�ŐV��DM��`�t�@�C����������擾���āAD-Tools��`���|�W�g���֔��f����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub SynchronizeDMCommandButton_Click()
   
    '���͏���ۑ�����
    '�e�[�u����`��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value = DTools.TableCheckBox.value
    'DM��`���ʕ��i�[�ꏊ
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value = DTools.LatestDMDefineFileTextBox.Text
    
    Dim dir As String
    dir = DTools.LatestDMDefineFileTextBox.Text
    
    If dir = "" Then
        MsgBox "D-Tools��`���|�W�g����ݒ肵�Ă��������B"
        Exit Sub
    End If
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools��`���|�W�g�����w�肵�Ă��������B"
        Exit Sub
    End If
    
    Dim objFSO As New FileSystemObject
    
    If objFSO.FileExists(DATA_SOURCE_DIR) = False Then
        MsgBox "�w���D-Tools��`���|�W�g�����s���݁I"
        Exit Sub
    End If
    
    On Error GoTo result
    
    
    'D-Tools��`���|�W�g�����o�b�N�A�b�v���܂�
    Dim backupDMRepositoryDir As String
    backupDMRepositoryDir = Replace(DATA_SOURCE_DIR, objFSO.getFileName(DATA_SOURCE_DIR), "") & Format(Now, "yyyymmddHHMM") & "_" & objFSO.getFileName(DATA_SOURCE_DIR)
    objFSO.CopyFile DATA_SOURCE_DIR, backupDMRepositoryDir
    
    Dim dirFolder As Folder
    Set dirFolder = objFSO.GetFolder(dir)
    
    Dim result As Boolean
    result = SynchronizeDMDefineInfo(dirFolder)
      
result:
    If Err.Number <> 0 Then
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
        Call ShowErrorMsg("SynchronizeDMCommandButton_Click")
    ElseIf result = False Then
        'D-Tools��ʂ��N���[�Y����
        'Call CloseForm
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
        '�������b�Z�[�W��񎦂���
        MsgBox "�����������܂����B"
    End If
        
End Sub
'***********************************************************************************************************************
' �@�\   : �e�[�u����`�����W�񂷂�
' �T�v   :
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub ExcleExtracteCommandButton_Click()
       
    '���͏���ۑ�����
    '�e�[�u����`��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value = DTools.TableCheckBox.value
    'DM��`���ʕ��i�[�ꏊ
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value = DTools.LatestDMDefineFileTextBox.Text
    
    Dim dir As String
    dir = DTools.LatestDMDefineFileTextBox.Text
    
    If dir = "" Then
        MsgBox "D-Tools��`���|�W�g����ݒ肵�Ă��������B"
        Exit Sub
    End If
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools��`���|�W�g�����w�肵�Ă��������B"
        Exit Sub
    End If
    
    Dim objFSO As New FileSystemObject
    
    If objFSO.FileExists(DATA_SOURCE_DIR) = False Then
        MsgBox "�w���D-Tools��`���|�W�g�����s���݁I"
        Exit Sub
    End If
    
    'On Error GoTo result
    
    Dim tableListSheetName As String
    tableListSheetName = "�ڎ�"
    
    '���s���ʂ̏o�͐�
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '���ʏo�͗p�V�[�g���쐬����
    Call createNewSheet(tableListSheetName, resultWorkBook)
    
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 1) = "��"
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 2) = "�V�[�g��"
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 3) = "�G�N�Z���t�@�C����"
    
    Dim dirFolder As Folder
    Set dirFolder = objFSO.GetFolder(dir)
    
    Dim result As Boolean
    'result = ExcleExtracte_sheetCopy(dirFolder)
    result = ExcleExtracteForDic(dirFolder)
    
    fileCounter = OPERATION_WORKBOOK.Sheets(tableListSheetName).UsedRange.Rows.Count
    '�r����t����
     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 3).Address).columns.AutoFit
    
result:
    If Err.Number <> 0 Then
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
        Call ShowErrorMsg("ExcleExtracteCommandButton_Click")
    ElseIf result = False Then
        'D-Tools��ʂ��N���[�Y����
        'Call CloseForm
    Else
        'D-Tools��ʂ��N���[�Y����
        Call CloseForm
        '�������b�Z�[�W��񎦂���
        MsgBox "�W�񊮗����܂����B"
    End If
End Sub

'***********************************************************************************************************************
' �@�\   : �_�������o�^�@�\
' �T�v   : �u�o�^�V�[�g���쐬����v�{�^�����������鎞�A�_�������o�^�p�V�[�g���쐬���āA�Y�V�[�g���ŁA�����o�^���s��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateRegisterSheetCommandButton_Click()
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools��`���|�W�g����ݒ肵�Ă��������B"
        Exit Sub
    End If
    
    '�����o�^�p�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "�_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "���l"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 4) = "�ǉ���"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 5) = "�ǉ���"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 6) = "�폜�t���O"

    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:C1").ColumnWidth = 15
    '�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("H1").Left, _
                                 Range("H1").Top, _
                                 Range("H1").Width, _
                                 Range("H1").Height)
        .OnAction = "SearchFromDic"
        .Characters.Text = "��������"
    End With
    
    '�o�^�{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("J1").Left, _
                                 Range("J1").Top, _
                                 Range("J1").Width, _
                                 Range("J1").Height)
        .OnAction = "RegisterToDic"
        .Characters.Text = "�����o�^"
    End With
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' �@�\   : �_���ϊ��@�\
' �T�v   : �u�_���ϊ��p�V�[�g���쐬����v�{�^�����������鎞�A�_���ϊ��p�V�[�g���쐬���āA�Y�V�[�g���ŁA�_���ϊ����s��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub LogicalToPhysicalCommandButton_Click()
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools��`���|�W�g����ݒ肵�Ă��������B"
        Exit Sub
    End If
    
    '�����o�^�p�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "�_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "���l"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 40
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "LogicalToPhysicalByDic"
        .Characters.Text = "�ϊ�����"
    End With
    
    
    '�ϊ��ł��Ȃ����A�񎦗p�`�F�b�N�{�b�N�X�����쐬����
    With ActiveSheet.CheckBoxes.Add(Range("G1").Left, _
                                   Range("G1").Top, _
                                   Range("G1").Width * 4, _
                                   Range("G1").Height)
        .Characters.Text = "�������ϊ��ł��Ȃ��ꍇ�A�����ϊ��̌��ʂ�񎦂���"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : �t�@�C���}�[�W�@�\
' �T�v   : �}�[�W��f�B���N�g���z���ɂ��ׂẴt�@�C�����}�[�W���f�B���N�g���z���ɒT���āA�}�[�W��f�B���N�g���փR�s�[����B
'          �}�[�W���ɕ���������ꍇ�A�ŏ����������t�@�C�����}�[�W��f�B���N�g���ɃR�s�[����B
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub MergeCommandButton_Click()
        
End Sub

'***********************************************************************************************************************
' �@�\   : �Z���������F�t���@�\
' �T�v   : �w��Z���ɁA�w�蕶�����Ŏw��F��t����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub SetColorInCellCommandButton_Click()
    '���͏���ۑ�����
    '�J�n�Z��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(24, 3).value = DTools.SetColorInCellStartTextBox.Text
    '�I���Z��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(25, 3).value = DTools.SetColorInCellEndTextBox.Text
    '�J�n����
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(26, 3).value = DTools.SetColorInCellCharStartTextBox.Text
    '�F�t��������
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(27, 3).value = DTools.SetColorInCellCharLengTextBox.Text
    '��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(28, 3).value = DTools.SetColorInCellRedColorOptionButton.value
    '��
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(29, 3).value = DTools.SetColorInCellBlueColorOptionButton.value
    
    Dim colorIndex As Integer
    
    If DTools.SetColorInCellRedColorOptionButton.value = xlOn Then
        colorIndex = 3
    ElseIf DTools.SetColorInCellBlueColorOptionButton.value = xlOn Then
        colorIndex = 2
    Else
        colorIndex = 3
    End If
        
    Dim setColorCell As Range
    For Each setColorCell In ActiveSheet.Range(DTools.SetColorInCellStartTextBox.Text & ":" & DTools.SetColorInCellEndTextBox.Text)
        Dim startIndex As Integer
        startIndex = InStr(setColorCell.value, DTools.SetColorInCellCharStartTextBox.Text)
        setColorCell.Characters(Start:=startIndex, Length:=Val(DTools.SetColorInCellCharLengTextBox.Text)).Font.colorIndex = colorIndex
    Next
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
    
End Sub
'***********************************************************************************************************************
' �@�\   : DB�J��������DTO���ɕύX���鏈��
' �T�v   : DB�����J��������DTO�p�̕ϐ����ɕύX����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub DbColumnNameChangeToDtoNameCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DB�J����������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DTO�ϐ���"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "ChangeDbColumnNameToDTOVar"
        .Characters.Text = "�ϊ�����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : DB�J���������DTO�N���X���쐬����
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateDtoClassCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DB�J�����_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DB�J����������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "DB�J�����̌^"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateDTOByDB"
        .Characters.Text = "�쐬����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : SqlMap�p�擾���ڃ}�b�s���O�쐬
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateSqlMapConfigCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DB�J�����_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DB�J����������"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateSqlMap"
        .Characters.Text = "�쐬����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : SqlMap�p�擾���ڃ}�b�s���O�쐬
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateCSVDtoClassCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DB�J�����_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DB�J����������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "DB�J�����̌^"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateCSVDTO"
        .Characters.Text = "�쐬����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : ���[�U�[��`�����DTO�N���X���쐬����
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CreateDtoClassByUserCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "�ϐ��_����"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "�ϐ�������"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "�ϐ��^"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateDTOByUser"
        .Characters.Text = "�쐬����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : ���[�U�[��`�����DTO�N���X���쐬����
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub ExtractFilesToolCommandButton_Click()
    '��Ɨp�V�[�g���쐬����
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "���[�gDIR"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "�t�@�C���p�X"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "���o����"


    '�r����t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '��̕���������
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '�F��t����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '�s�̍����𒲐�����
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 80
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '�ϊ�����{�^�����쐬����
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "ExtractFiles"
        .Characters.Text = "���o����"
    End With
    
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
'***********************************************************************************************************************
' �@�\   : COBOL�̉����ԍ����ڍא݌v���̋L�q�R�ꂪ�Ȃ������`�F�b�N����
' �T�v   : �Ȃ�
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Private Sub CheckSSCommandButton_Click()
    '�݌v���`�F�b�N�@�\������������
    'Call AddMenuca
    Call �݌v���`�F�b�N_�蓮�I��
    
    'D-Tools��ʂ��N���[�Y����
    Call CloseForm
End Sub
