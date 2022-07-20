Attribute VB_Name = "Utils"
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �A�h�C���ɑ���f�[�^���i�[���邽�߂̃V�[�g�����쐬����B���������̂݁B
' ����   : String �V�[�g���AWorkbook�@�ǉ��Ώۂ̃G�N�Z��
' �߂�l : �V�V�[�g��
'***********************************************************************************************************************
Public Function createNewSheet(resultSheetName As String, addWorkbook As Workbook) As String
    Dim existFlag As Boolean
    existFlag = checkSheetNameExist(resultSheetName, addWorkbook)
       
    '�V�V�[�g�쐬
    If existFlag = False Then
        addWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).name = resultSheetName
        '�\���`���𕶎���ɐݒ肷��
        addWorkbook.Worksheets(resultSheetName).Cells.NumberFormatLocal = "@"
    End If
    
    createNewSheet = addWorkbook.Worksheets(Worksheets.Count).name
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �����̃V�[�g���A�V�[�g�����A�V�����V�[�g��ǉ�����B
' ����   : String �V�[�g���AWorkbook�@�ǉ��Ώۂ̃G�N�Z��
' �߂�l : ��
'***********************************************************************************************************************
Public Function outPutResultSheet(resultSheetName As String, addWorkbook As Workbook)
    Dim existFlag As Boolean
    existFlag = False
    
    Dim sheetsCount As Integer
    sheetsCount = addWorkbook.Worksheets.Count
    
    existFlag = checkSheetNameExist(resultSheetName & sheetsCount + 1, addWorkbook)
    
    '�V�V�[�g�쐬
    If existFlag = False Then
        resultSheetName = resultSheetName & sheetsCount + 1
        addWorkbook.Sheets.Add(After:=Worksheets(sheetsCount)).name = resultSheetName
    Else
        resultSheetName = resultSheetName & sheetsCount + 1
        existFlag = checkSheetNameExist(resultSheetName, addWorkbook)
        If existFlag = True Then
            Do
                Dim i As Integer
                i = 1
                resultSheetName = resultSheetName & "(" & i & ")"
                existFlag = checkSheetNameExist(resultSheetName, addWorkbook)
                i = i + 1
            Loop Until existFlag = False
        End If
        addWorkbook.Sheets.Add(After:=Worksheets(sheetsCount)).name = resultSheetName
    End If
    '�\���`���𕶎���ɐݒ肷��
    addWorkbook.Worksheets(resultSheetName).Cells.NumberFormatLocal = "@"
    '�o�͌��ʂ̃V�[�g��
    RESULT_SHEET_NAME = resultSheetName
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��V�[�g�����w��G�N�Z���ɑ��݂��邱�Ƃ��`�F�b�N����
' ����   : String �V�[�g���AWorkbook�@�w��̃G�N�Z��
' �߂�l : �`�F�b�N���ʁiTRUE:���݁^FALSE�F�s���݁j
'***********************************************************************************************************************
Public Function checkSheetNameExist(resultSheetName As String, addWorkbook As Workbook) As Boolean
    Dim existFlag As Boolean
    existFlag = False
    '�V�[�g���݃`�F�b�N
    For i = 1 To addWorkbook.Sheets.Count
        If resultSheetName = addWorkbook.Sheets(i).name Then
            existFlag = True
            Exit For
        End If
    Next
checkSheetNameExist = existFlag
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �S�̕ϐ�����A�ڑ����𗘗p���āADB�ڑ�����
' ����   : �Ȃ�
' �߂�l : DB�ڑ��̃I�u�W�F�N�g
'***********************************************************************************************************************
Public Function connDB() As ADODB.Connection
    On Error GoTo errorHandler
    Dim ADOConnection As New ADODB.Connection
    ADOConnection.Open DB_CONN_INFO_STR

errorHandler:
    If Err.Number <> 0 Then
        'Call ShowErrorMsg("connDB")
    End If
    
    Set connDB = ADOConnection
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �S�̕ϐ�����A�ڑ����𗘗p���āAAccessDB�ڑ�����
' ����   : �Ȃ�
' �߂�l : DB�ڑ��̃I�u�W�F�N�g
'***********************************************************************************************************************
Public Function connAccessDB() As ADODB.Connection
    On Error GoTo errorHandler
    Dim ADOConnection As New ADODB.Connection
    Dim dataSource As String
    dataSource = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATA_SOURCE_DIR & ";"
    
    ADOConnection.Open dataSource

errorHandler:
    If Err.Number <> 0 Then
        'Call ShowErrorMsg("connAccessDB")
    End If
    
    Set connAccessDB = ADOConnection

End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : DB����e�[�u���̕��������A�e�[�u���̕����A�_���J���������擾����
' ����   : String�@�e�[�u���̕�����
' �߂�l : �J�����������Ƙ_�����̌��ʏW��
'***********************************************************************************************************************
Function GetTableColumnsNameFromDB(tableNameEN As String) As Collection
    Dim sql As String
    
    'sql = "SELECT COMMENTS, COLUMN_NAME FROM USER_COL_COMMENTS WHERE TABLE_NAME = '" & UCase(tableNameEN) & "'"
    
    sql = "SELECT " & _
                "T1.COMMENTS AS COMMENTS," & _
                "T1.COLUMN_NAME AS COLUMN_NAME," & _
                "T2.DATA_TYPE AS DATA_TYPE," & _
                "T2.DATA_LENGTH AS DATA_LENGTH," & _
                "T2.NULLABLE AS NULLABLE," & _
                "T4.CONSTRAINT_TYPE AS CONSTRAINT_TYPE " & _
            "FROM " & _
                "USER_COL_COMMENTS T1," & _
                "USER_TAB_COLUMNS T2," & _
                "USER_CONS_COLUMNS T3," & _
                "USER_CONSTRAINTS T4 " & _
            "WHERE " & _
                "T1.TABLE_NAME = T2.TABLE_NAME AND " & _
                "T2.TABLE_NAME = T3.TABLE_NAME AND " & _
                "T2.COLUMN_NAME = T3.COLUMN_NAME AND " & _
                "T3.CONSTRAINT_NAME = T4.CONSTRAINT_NAME AND " & _
                "T1.TABLE_NAME = '" & UCase(tableNameEN) & "'"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connDB()
    
    Dim ADORecordset As New ADODB.recordset
    ADORecordset.Open sql, ADOConnection
    
    Dim columnNameCollection As New Collection
    
    Do Until ADORecordset.EOF
        Dim columnName() As String
        ReDim columnName(5)
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("COMMENTS").value) Then
            columnName(0) = ""
        Else
            columnName(0) = resultFields("COMMENTS").value
        End If
        
        columnName(1) = resultFields("COLUMN_NAME").value
        columnName(2) = resultFields("DATA_TYPE").value
        columnName(3) = resultFields("DATA_LENGTH").value
        
        If IsNull(resultFields("NULLABLE").value) Then
            columnName(4) = ""
        Else
                If resultFields("NULLABLE").value = "N" Then
                    columnName(4) = "NULL�s��"
                Else
                    columnName(4) = ""
                End If
        End If
        
        If resultFields("CONSTRAINT_TYPE").value = "P" Or resultFields("CONSTRAINT_TYPE").value = "U" Then
            columnName(5) = "1"
        End If
        columnNameCollection.Add (columnName)
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    Set GetTableColumnsNameFromDB = columnNameCollection
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : DM��`���|�W�g������e�[�u���̕��������A�e�[�u���̕����A�_���J���������擾����
' ����   : String�@�e�[�u���̕�����
' �߂�l : �J�����������Ƙ_�����̌��ʏW��
'***********************************************************************************************************************
Public Function GetTableColumnsNameFromDMRepository(tableNameEN As String) As Collection
    Dim sql As String
    sql = "SELECT [������_�a��],[�J������_�p��],[�f�[�^�^],[����],[��L�[],[NULL] FROM [������`��] WHERE [�e�[�u����_�p��] = '" & UCase(tableNameEN) & "' ORDER BY [No]"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Dim ADORecordset As New ADODB.recordset
    Set ADORecordset = ADOConnection.Execute(sql)
    
    Dim columnNameCollection As New Collection
    
    '0���`�F�b�N
    Do Until ADORecordset.EOF
        Dim columnName() As String
        ReDim columnName(5)
        Set resultFields = ADORecordset.Fields
        columnName(0) = resultFields("������_�a��").value
        columnName(1) = resultFields("�J������_�p��").value
        columnName(2) = resultFields("�f�[�^�^").value
        columnName(3) = resultFields("����").value
        columnName(4) = resultFields("NULL").value
        columnName(5) = resultFields("��L�[").value
        columnNameCollection.Add (columnName)
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    '�擾���ʂ�ԋp����
    Set GetTableColumnsNameFromDMRepository = columnNameCollection
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �e�[�u���������̏ꍇ�ADM��`���|�W�g������e�[�u���_�����擾����B�e�[�u���_�����̏ꍇ�ADM��`���|�W�g������e�[�u�����������擾����B
' ����   : String�@�e�[�u����
' �߂�l : Collection �e�[�u���_�����ƃe�[�u���������̏W��
'***********************************************************************************************************************
Public Function GetTableNameFromDMRepository(tableName As String) As Collection
    Dim tableNameInfoCol As New Collection
    Dim tableNameInfo() As String
    Dim tableNameEN As String
    Dim tableNameJP As String
    
    If tableName = "" Or tableName = Null Then
        GoTo existFunction
    End If
    
    Dim sql As String
    If IsContainJapanese(tableName) Then
        tableName = Replace(tableName, "*", "%")
        sql = "SELECT [�e�[�u����_�a��], [�e�[�u����_�p��] FROM [�e�[�u����`��] WHERE [�e�[�u����_�a��] LIKE '" & tableName & "'"
    Else
        tableName = Replace(tableName, "*", "%")
        sql = "SELECT [�e�[�u����_�a��], [�e�[�u����_�p��] FROM [�e�[�u����`��] WHERE [�e�[�u����_�p��] LIKE '" & UCase(tableName) & "'"
    End If
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Dim ADORecordset As New ADODB.recordset
    Set ADORecordset = ADOConnection.Execute(sql)
    
    Do Until ADORecordset.EOF
        ReDim tableNameInfo(1)
        Set resultFields = ADORecordset.Fields
        tableNameEN = resultFields("�e�[�u����_�p��").value
        tableNameJP = resultFields("�e�[�u����_�a��").value
        
        tableNameInfo(0) = tableNameEN
        tableNameInfo(1) = tableNameJP
        
        tableNameInfoCol.Add (tableNameInfo)
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
existFunction:
    
    Set GetTableNameFromDMRepository = tableNameInfoCol
    
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �e�[�u���������̏ꍇ�ADB����e�[�u���_�����擾����B�e�[�u���_�����̏ꍇ�ADB����e�[�u�����������擾����B
' ����   : String�@�e�[�u����
' �߂�l : Collection �e�[�u���_�����ƃe�[�u���������̏W��
'***********************************************************************************************************************
Public Function GetTableNameFromDB(tableName As String) As Collection
    
    Dim tableNameInfoCol As New Collection
    Dim tableNameInfo() As String
    Dim tableNameEN As String
    Dim tableNameJP As String
    
    If tableName = "" Or tableName = Null Then
        GoTo existFunction
    End If

    Dim sql As String
    If StrConv(Application.GetPhonetic(Replace(UCase(tableName), "ID", "")), vbHiragana) Like "*[��-��]*" Then
        sql = "SELECT NVL(COMMENTS,'') AS COMMENTS, NVL(TABLE_NAME,'') AS TABLE_NAME FROM USER_TAB_COMMENTS WHERE COMMENTS like '%" & tableName & "%'"
    Else
        sql = "SELECT NVL(COMMENTS,'') AS COMMENTS, NVL(TABLE_NAME,'') AS TABLE_NAME FROM USER_TAB_COMMENTS WHERE TABLE_NAME like '%" & UCase(tableName) & "%'"
    End If
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connDB()
       
    Dim rs As New ADODB.recordset
    rs.Open sql, ADOConnection

    Do Until rs.EOF
        ReDim tableNameInfo(1)
        Set resultFields = rs.Fields
        
        tableNameEN = resultFields("TABLE_NAME").value
        
        If IsNull(resultFields("COMMENTS")) Then
            tableNameJP = ""
        Else
            tableNameJP = resultFields("COMMENTS").value
        End If
            
        tableNameInfo(0) = tableNameEN
        tableNameInfo(1) = tableNameJP
        
        tableNameInfoCol.Add (tableNameInfo)
        
        rs.MoveNext
    Loop
    rs.Close
    ADOConnection.Close
    
existFunction:
    Set GetTableNameFromDB = tableNameInfoCol
    
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : DB�ɂ���̂��ׂẴ��[�U�[�̃e�[�u���̕��������擾����
' ����   : ��
' �߂�l : �e�[�u���������̏W��
'***********************************************************************************************************************
Public Function GetAllTableNameEN() As Collection
    Dim sql As String
    Dim resutList As New Collection
    sql = "SELECT �e�[�u����_�p�� AS TABLE_NAME FROM �e�[�u����`��"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    If ADOConnection.State = 0 Then
        Set ADOConnection = connDB()
        sql = "SELECT TABLE_NAME FROM USER_TABLES"
    End If
       
    Dim rs As New ADODB.recordset
    rs.Open sql, ADOConnection
    
    Do Until rs.EOF
        Set resultFields = rs.Fields
        Call resutList.Add(resultFields("TABLE_NAME").value)
        rs.MoveNext
    Loop
    
    
    rs.Close
    ADOConnection.Close
    
    '�擾���ʂ�ԋp����
    Set GetAllTableNameEN = resutList
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��s�ɁA�e�[�u���������𑶍݂��Ă��邱�Ƃ��`�F�b�N����
' ����   : Integer�@�s�ԍ�
' �߂�l : �`�F�b�N����(TRUE:����/FALSE�F�s����)
'***********************************************************************************************************************
Public Function checkTableNameExistInRow(rowNo As Integer) As Boolean
    '�e�[�u����������T��
    Dim tableNameEN As String
    Dim rowRange As Range
    
    Dim allTableNameEN As New Collection
    Set allTableNameEN = GetAllTableNameEN
    
    Dim existFlag As Boolean
    
    For Each rowRange In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
        existFlag = False
        If StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[��-��]*" Then
            '���{����܂ޏꍇ�A���s���̎��̃Z�����`�F�b�N����B
            GoTo Continue
        Else
            tableNameEN = UCase(rowRange.value)
            For i = 1 To allTableNameEN.Count
                If tableNameEN = allTableNameEN(i) Then
                    existFlag = True
                    Exit For
                End If
                
            Next
        End If
        
        If existFlag = True Then
            Exit For
        End If
Continue:
    Next
    
    checkTableNameExistInRow = existFlag
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��s�ɁA�w�蕶���i�����j�𑶍݂��Ă��邱�Ƃ��`�F�b�N����
' ����   : Integer�@�s�ԍ��AString �w��̕����z��
' �߂�l : �`�F�b�N����(0:�s���� 1:���݂��� 2�F�S������)
'***********************************************************************************************************************
Public Function checkStrsExistInRow(rowNo As Integer, strs As Variant) As String
    Dim str As Variant
    Dim rowRange As Range
    Dim cellOBJ As Range
    Dim cellValue As String
    
    Dim checkResult As String
    Dim existCounter As Integer
    existCounter = 0
    For Each str In strs
        For Each cellOBJ In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
            cellValue = cellOBJ.value
            If str = cellValue Then
                existCounter = existCounter + 1
                Exit For
            End If
        Next
    Next

    If existCounter = 0 Then
        checkResult = "0"
    ElseIf existCounter < UBound(strs) Then
        checkResult = "1"
    Else
        checkResult = "2"
    End If
        
    checkStrsExistInRow = checkResult
    
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��e�[�u���������̃J���������擾����
' ����   : String�@�e�[�u���̕�����
' �߂�l : �J�������Ƃ̃J�����������ANULL�ۂ̔z��̏W��
'***********************************************************************************************************************
Public Function GetTabColumns(tableNameEN As String) As Collection
    Dim sql As String
    Dim resutList As New Collection
    sql = "SELECT [�J������_�p��] AS COLUMN_NAME,IIF([��L�[] <> '','Y','N') AS ISKEY FROM [������`��] WHERE [�e�[�u����_�p��] = '" & UCase(tableNameEN) & "'"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connAccessDB()
    If ADOConnection.State = 0 Then
        Set ADOConnection = Utils.connDB()
        sql = "SELECT " & _
                    "T2.COLUMN_NAME AS COLUMN_NAME," & _
                    "CASE WHEN T4.CONSTRAINT_TYPE = 'U' THEN 'Y' WHEN T4.CONSTRAINT_TYPE = 'P' THEN 'Y' ELSE 'N' END AS ISKEY " & _
                "FROM " & _
                    "USER_TAB_COLUMNS T2," & _
                    "USER_CONS_COLUMNS T3," & _
                    "USER_CONSTRAINTS T4 " & _
                "WHERE " & _
                    "T2.TABLE_NAME = T3.TABLE_NAME AND " & _
                    "T2.COLUMN_NAME = T3.COLUMN_NAME AND " & _
                    "T3.CONSTRAINT_NAME = T4.CONSTRAINT_NAME AND " & _
                    "T2.TABLE_NAME = '" & UCase(tableNameEN) & "'"
    End If

    Dim rs As New ADODB.recordset
    rs.Open sql, ADOConnection
    
    Do Until rs.EOF
        Set resultFields = rs.Fields
        Dim data() As String
        ReDim data(1)
        data(0) = resultFields("COLUMN_NAME").value
        data(1) = resultFields("ISKEY").value
        resutList.Add (data)
        rs.MoveNext
    Loop
    
    rs.Close
    ADOConnection.Close
    
    '�擾���ʂ�ԋp����
    Set GetTabColumns = resutList

End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : DB����w��e�[�u���������̃J���������擾����
' ����   : String�@�e�[�u���̕�����
' �߂�l : �J�������Ƃ̃J�����������ANULL�ہA�^�A�T�C�Y�̔z��̏W��
'***********************************************************************************************************************
Public Function GetTabColumnInfoFromDB(tableNameEN As String) As Collection
    Dim sql As String
    Dim resutList As New Collection
    sql = "SELECT COLUMN_NAME,NULLABLE, DATA_TYPE,DATA_LENGTH FROM USER_TAB_COLUMNS WHERE TABLE_NAME = '" & UCase(tableNameEN) & "'"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connDB()
       
    Dim rs As New ADODB.recordset
    rs.Open sql, ADOConnection
    
    Do Until rs.EOF
        Set resultFields = rs.Fields
        Dim data() As String
        ReDim data(3)
        data(0) = resultFields("COLUMN_NAME").value
        data(1) = resultFields("NULLABLE").value
        data(3) = resultFields("DATA_TYPE").value
        data(4) = resultFields("DATA_LENGTH").value
        resutList.Add (data)
        rs.MoveNext
    Loop
    
    rs.Close
    ADOConnection.Close
    
    '�擾���ʂ�ԋp����
    Set GetTabColumnInfo = resutList
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ��������̎w��̕����񂩂�A�w��̕�����܂ł̕������؂���
' ����   : String �؎�Ώۂ̕�����AString �؎�J�n������AString �؎�I��������
' �߂�l : �؎敶����
'***********************************************************************************************************************
Public Function GetStrFromStr(findStr As String, startStr As String, endstr As String) As String
    
    If findStr = "" Then
        GetStrFromStr = ""
        Exit Function
    End If
    
    If startStr <> "" Then
        startIndex = InStr(1, findStr, startStr, vbTextCompare) + Len(startStr)
    Else
        startIndex = 1
    End If
    
    If endstr <> "" Then
        endIndex = InStr(1, findStr, endstr, vbTextCompare) - 1
    Else
        endstr = Len(findStr)
    End If
    
    
    GetStrFromStr = Mid(findStr, startIndex, endIndex - startIndex)
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ���������space�A�^�u�R�[�h�A���s�R�[�h���󔒂ɒu������
' ����   : �u���Ώۂ̕�����
' �߂�l : �u����̕�����
'***********************************************************************************************************************
Public Function ReplaceSTNToNull(str As String) As String
    '��
    str = Replace(str, Chr(32), "")
    '�^�u
    str = Replace(str, Chr(9), "")
    '���sLF
    str = Replace(str, Chr(10), "")
    '���sCR
    str = Replace(str, Chr(13), "")
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ��������̃^�u�R�[�h�A���s�R�[�h���󔒂ɒu������
' ����   : �u���Ώۂ̕�����
' �߂�l : �u����̕�����
'***********************************************************************************************************************
Public Function ReplaceTNToSpace(str As String) As String
    '�^�u
    str = Replace(str, Chr(9), " ")
    '���sLF
    str = Replace(str, Chr(10), " ")
    '���sCR
    str = Replace(str, Chr(13), " ")
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��s���󔒂̍s���ǂ������`�F�b�N����
' ����   : Integer �s�ԍ�
' �߂�l : �`�F�b�N���ʁiTRUE�FAll�Z�����󔒁^FALSE:�󔒈ȊO�̃Z���𑶍݂���j
'***********************************************************************************************************************
Public Function RowIsAllSpace(rowIndex As Integer) As Boolean
    Dim counter As Integer
    counter = 0
    On Error GoTo errorHander
    counter = ActiveSheet.Rows(rowIndex).SpecialCells(xlCellTypeConstants).Count
    If counter <> 0 Then
        RowIsAllSpace = False
    Else
        RowIsAllSpace = True
    End If
    
errorHander:
    If Err.Number <> 0 Then
        RowIsAllSpace = True
    End If
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��s�̈�ڋ󔒂ł͂Ȃ��Z�����擾����
' ����   : Integer �s�ԍ�
' �߂�l : Range �Z���I�u�W�F�N�g
'***********************************************************************************************************************
Public Function GetNotNullCellInOneRow(rowIndex As Integer) As Range
    Dim firstCell As Range
    
    On Error GoTo errorHander
    For Each firstCell In ActiveSheet.Rows(rowIndex).SpecialCells(xlCellTypeConstants)
        Exit For
    Next
    
errorHander:
    If Err.Number <> 0 Then
        Exit Function
    End If
    
    Set GetNotNullCellInOneRow = firstCell
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��Z�����̓���ɁA��ڋ󔒂ł͂Ȃ��Z�����擾����
' ����   : Range �Z���I�u�W�F�N�g
' �߂�l : Range �Z���I�u�W�F�N�g
'***********************************************************************************************************************
Public Function GetNotNullCellUnderOneCell(fromRange As Range) As Range
    Dim endCell As Range
  
    On Error GoTo errorHander
    'For Each endCell In ActiveSheet.Range(Replace(ActiveSheet.Cells(fromRange.Row + 1, fromRange.column).Address & ":" & ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, fromRange.column).Address, "$", "")).SpecialCells(xlCellTypeConstants)
    For Each endCell In ActiveSheet.Range(Replace(ActiveSheet.Cells(fromRange.Row + 1, 1).Address & ":" & ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, fromRange.Column).Address, "$", "")).SpecialCells(xlCellTypeConstants)
        Exit For
    Next
    
errorHander:
    If Err.Number <> 0 Then
        'Call ShowErrorMsg("GetNotNullCellUnderOneCell")
        Set endCell = ActiveSheet.Range(ActiveSheet.Cells(ActiveSheet.UsedRange.Rows.Count, fromRange.Column).Address)
    End If
    
    Set GetNotNullCellUnderOneCell = ActiveSheet.Range(ActiveSheet.Cells(endCell.Row, fromRange.Column).Address)
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��J�n�s����w��I���s�܂ŁA�O���[�v������
' ����   : Integer �J�n�s�@Integer�@�I���s
' �߂�l : ��
'***********************************************************************************************************************
Function Group(groupStartIndex As Integer, groupEndIndex As Integer)
    On Error GoTo Continue
    ActiveSheet.Rows(groupStartIndex & ":" & groupEndIndex).Group
Continue:
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �w��Z���̓��e���A�O���[�v�����߂̊J�n�_���ǂ����𔻒f����
' ����   : String �Z���̓��e
' �߂�l : Boolean �`�F�b�N����
'***********************************************************************************************************************
Function IsGroupStartRow(cellValue As String) As Boolean
    Dim checkResult As Boolean
    checkResult = False
    
    If cellValue Like "��*" Then
        checkResult = True
    ElseIf cellValue Like "[(][0-9]*" Then
        checkResult = True
    ElseIf cellValue Like "[�i][0-9]*" Then
        checkResult = True
    ElseIf cellValue Like "[0-9]-*" Then
        checkResult = True
    ElseIf cellValue Like "[0-9].*" Then
        checkResult = True
    Else
        checkResult = False
    End If

    IsGroupStartRow = checkResult
End Function
'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �w��s����e�[�u����������T���āA�ԋp����
' ����   : Integer �s�ԍ�
' �߂�l : ���������e�[�u��������
'***********************************************************************************************************************
Public Function CreateSql_GetTableName(rowNo As Integer) As String
    '�e�[�u����������T��
    Dim tableNameEN As String
    Dim rowRange As Range
    For Each rowRange In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
        If rowRange.value = "" Then
            '�󔒂̏ꍇ�A���s���̎��̃Z�����`�F�b�N����B
            GoTo Continue
        ElseIf StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[��-��]*" Then
            '���{����܂ޏꍇ�A���s���̎��̃Z�����`�F�b�N����B
            '�e�[�u���������Ƒz��
            PUB_TEMP_VAR_STR = rowRange.value
            GoTo Continue
        Else
            tableNameEN = UCase(rowRange.value)
            Exit For
        End If
        
Continue:
    Next
    
    CreateSql_GetTableName = tableNameEN
    
End Function

'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �w��s���J�����s�̏ꍇ�A�J�������̏W���A�J�����̊J�n�C���f�b�N�X�A�J�����̏I���C���f�b�N�X
' ����   : Integer �s�ԍ�
' �߂�l : �ԋp���������ʂ̔z��
'***********************************************************************************************************************
Public Function CreateSql_GetTableColumns(rowNo As Integer) As Variant
    
    Dim tableColumns() As String
    Dim columnIndex As Integer
    Dim columnEndIndex As Integer
    Dim columnStartIndex As Integer
    
    Dim result As Variant
    
    columnIndex = 0
    ReDim tableColumns(ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants).Count - 1)
    
    For Each rowRange In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
                
                If StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[��-��]*" Then
                    '���{����܂ޏꍇ�A�����J�����ł͂Ȃ�
                    Erase tableColumns
                    columnStartIndex = 0
                    columnEndIndex = 0

                    Exit For
                End If
                
                tableColumns(columnIndex) = UCase(rowRange.value)
                
                columnIndex = columnIndex + 1
                
                If columnIndex = 1 Then
                    columnStartIndex = rowRange.Column
                End If
    Next
    
    columnEndIndex = columnStartIndex + ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants).Count - 1
    ReDim result(2)
    result(0) = tableColumns
    result(1) = columnStartIndex
    result(2) = columnEndIndex
    
    CreateSql_GetTableColumns = result
End Function
    
'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �w��s�̎w���̃f�[�^��z��Ɏ������āA�ԋp����
' ����   : Integer �s�ԍ�, columnStartIndex �J�����J�n�C���f�b�N�X, columnEndIndex �J�����I���C���f�b�N�X
' �߂�l : �w��s�̎w���̃f�[�^�z��
'***********************************************************************************************************************
Public Function CreateSql_GetData(rowNo As Integer, columnStartIndex As Integer, columnEndIndex As Integer) As Variant
    Dim data() As String
    ReDim data(columnEndIndex - columnStartIndex)
    Dim i As Integer
    i = 0
    Do While i <= columnEndIndex - columnStartIndex
        data(i) = ActiveSheet.Cells(rowNo, columnStartIndex + i).value
        i = i + 1
    Loop
    
    CreateSql_GetData = data

End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
' ����   : varArray  �z��
' �߂�l : ���茋�ʁi1:�z��/0:��̔z��/-1:�z�񂶂�Ȃ��j
'***********************************************************************************************************************
Public Function IsArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_

    If IsArray(varArray) Then
        IsArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        IsArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        IsArrayEx = 0
    End If
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ���݊J���Ă���t�@�C���̃f�B���N�g�����`�F�b�N���āA�󔒂̏ꍇ�A�f�B�X�N�g�b�v�̃f�B���N�g����ԋp����B
' ����   : �Ȃ�
' �߂�l : �f�B���N�g���̕�����
'***********************************************************************************************************************
Public Function GetSaveDir() As String
    Dim path As String
    path = ActiveWorkbook.path
    If path = "" Then
        Dim WSH As Variant
        Set WSH = CreateObject("WScript.Shell")
        path = WSH.SpecialFolders("Desktop")
        Set WSH = Nothing
    End If
    
    GetSaveDir = path
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : ���ʎq���擾����
' ����   : String ���ʎq�̎�ʁi�ݒ肵�Ȃ����A���ׂĂ̎��ʂ��擾����j
' �߂�l : collection�@���ʎq�R�[�h�Ǝ��ʎq���̏W��
'***********************************************************************************************************************
Public Function GetKeyValue(keyCode As String) As Collection
    Dim keyValueInfo() As String
    Dim keyValueInfoCollection As New Collection
    Dim sql As String
    Dim keyName As String
    Dim keyValue As String
    
    If keyCode = "1" Then
        sql = "SELECT [���ʎq�R�[�h], [���ʎq��] FROM [���ʎq�Ǘ��e�[�u��] WHERE [���ʎq���] = '1' AND [�폜�t���O] = '0'"
    ElseIf keyCode = "2" Then
        sql = "SELECT [���ʎq�R�[�h], [���ʎq��] FROM [���ʎq�Ǘ��e�[�u��] WHERE [���ʎq���] = '2' AND [�폜�t���O] = '0'"
    ElseIf keyCode = "" Then
        sql = "SELECT [���ʎq�R�[�h], [���ʎq��] FROM [���ʎq�Ǘ��e�[�u��] WHERE [�폜�t���O] = '0'"
    Else
        Exit Function
    End If
    
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Set ADORecordset = New ADODB.recordset
    ADORecordset.Open sql, ADOConnection
    
    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("���ʎq��")) Then
            keyName = ""
        Else
            keyName = resultFields("���ʎq��").value
        End If
        If IsNull(resultFields("���ʎq�R�[�h")) Then
            keyValue = ""
        Else
            keyValue = resultFields("���ʎq�R�[�h").value
        End If
        
        ReDim keyValueInfo(1)
        keyValueInfo(0) = keyName
        keyValueInfo(1) = keyValue
        
        keyValueInfoCollection.Add (keyValueInfo)
        
        ADORecordset.MoveNext
    Loop
    
    Set GetKeyValue = keyValueInfoCollection
    
End Function

'***********************************************************************************************************************
' �@�\   : �o�[�W�����`�F�b�N�@�\
' �T�v   : ���p���Ă���o�[�W�������ŐV���ǂ����`�F�b�N����
' ����   : �Ȃ�
' �߂�l : �Ȃ�
'***********************************************************************************************************************
Public Function CheckDdataVersion()
    On Error GoTo errorHandler

    'D-Tools��ʂ̏����ݒ�����{����
    Load DTools
    
    Dim ADOConnection As New ADODB.Connection
    Dim dataSource As String
    dataSource = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATA_SOURCE_DIR & ";"
    ADOConnection.Open dataSource
    
    '�ŐV�̃o�[�W��������񎦂���
    Dim sql As String
    Dim ADORecordset As New ADODB.recordset
    Dim �o�[�W����, ���C���e, �C����, �C����, �A�h�C���i�[�ꏊ As String
    Dim rowCount As Integer
    Dim usedVersion As String
    Dim result As Integer
    
    rowCount = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).UsedRange.Rows.Count
    
    Do Until usedVersion <> ""
        usedVersion = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).Cells(rowCount, 2).value
        rowCount = rowCount - 1
    Loop
            
    sql = "SELECT [�o�[�W����], [���C���e],[�C����],[�C����],[�A�h�C���i�[�ꏊ] FROM [Ddata�o�[�W�������] WHERE [ID] = (SELECT MAX([ID]) FROM [Ddata�o�[�W�������] WHERE [�o�[�W����] > " & Val(usedVersion) & ")"
    ADORecordset.Open sql, ADOConnection

    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        
        If IsNull(resultFields("�o�[�W����")) Then
            �o�[�W���� = ""
        Else
            �o�[�W���� = resultFields("�o�[�W����").value
        End If
        
        If IsNull(resultFields("���C���e")) Then
            ���C���e = ""
        Else
            ���C���e = resultFields("���C���e").value
        End If
        
        If IsNull(resultFields("�C����")) Then
            �C���� = ""
        Else
            �C���� = resultFields("�C����").value
        End If
        
        If IsNull(resultFields("�C����")) Then
            �C���� = ""
        Else
            �C���� = resultFields("�C����").value
        End If
        
        If IsNull(resultFields("�A�h�C���i�[�ꏊ")) Then
            �A�h�C���i�[�ꏊ = ""
        Else
            �A�h�C���i�[�ꏊ = resultFields("�A�h�C���i�[�ꏊ").value
        End If
          
        Dim versionMessage  As String
        
        versionMessage = "D-Tools�̍ŐV�o�[�W�����y" & �o�[�W���� & "�z�������[�X���܂����B" & vbCrLf & vbCrLf & _
                         "�C�����F" & �C���� & "  �C���ҁF" & �C���� & vbCrLf & vbCrLf & _
                         "�A�h�C���i�[�ꏊ�F" & vbCrLf & �A�h�C���i�[�ꏊ & vbCrLf & vbCrLf & _
                         "���C���e�F" & vbCrLf & ���C���e
                         
        versionMessage = versionMessage & vbCrLf & vbCrLf & vbCrLf & "�͂�(Y)����������ƁA�A�h�C���i�[�̏ꏊ���R�s�[���܂��B"
        versionMessage = versionMessage & vbCrLf & "������(N)����������ƁA����̍X�V���X�L�b�v���܂��B"
        
        result = MsgBox(versionMessage, vbYesNo + vbExclamation)
        
        If result = vbYes Then
            Dim myData As New DataObject
            myData.SetText (�A�h�C���i�[�ꏊ)
            myData.PutInClipboard
        ElseIf result = vbNo Then
            '���C�����V�[�g�ɁA���̍ŐV�̃o�[�W���������L�ڂ��āA����̍X�V�񎦂����Ȃ��Ȃ�܂��B
            rowCount = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).UsedRange.Rows.Count
            addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).Cells(rowCount + 1, 2).value = �o�[�W����
        End If
        
        ADORecordset.MoveNext
        
    Loop
    
errorHandler:
'D-Tools��ʂ̏����ݒ���������
    Call CloseForm
    
End Function
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �Ώە�������{�ꂪ���݂��邩���`�F�b�N
' ����   : �`�F�b�N�Ώە�����
' �߂�l : True�F���{�ꕶ�������݁AFalse�F���{�ꕶ�����s����
'***********************************************************************************************************************
Public Function IsContainJapanese(str As String) As Boolean
    Dim charStr As String
    For i = 1 To Len(str)
        charStr = Mid(str, i, 1)
        If StrConv(Application.GetPhonetic(charStr), vbHiragana) Like "*[��-��]*" Then
            IsContainJapanese = True
            Exit Function
        End If
    Next
    
    IsContainJapanese = False
End Function


'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �G���[�𔭐��������A�_�C�A���O�ɃG���[���e��\������B
' ����   : �Ȃ�
' �߂�l : �Ȃ�
'***********************************************************************************************************************
Public Function ShowErrorMsg(Optional functionName As String)
    Dim errorMsg As String
    
    errorMsg = errorMsg & "�G���[�ԍ��F" & Err.Number & vbNewLine
    errorMsg = errorMsg & "�G���[���e�F" & Err.Description & vbNewLine
    errorMsg = errorMsg & "�w���v�t�@�C�����F" & Err.HelpContext & vbNewLine
    errorMsg = errorMsg & "�v���W�F�N�g���F" & Err.Source & vbNewLine
    If functionName <> "" Then
        errorMsg = errorMsg & "���\�b�h���F" & functionName & vbNewLine
    End If
    Err.Clear
    
    '�ŏ���2���������A�G���[�������W����
    If Format(Now, "yyyymmdd") < "20990120" Then
        Call SaveErrorInfo(errorMsg)
    End If
    
    MsgBox errorMsg
    
End Function

'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �G���[�𔭐��������A�_�C�A���O�ɃG���[���e��\������B
' ����   : �Ȃ�
' �߂�l : �Ȃ�
'***********************************************************************************************************************
Public Function SaveErrorInfo(errorMsg As String)
    On Error GoTo errorHandler
    
    Dim ADOConnection As New ADODB.Connection
    Dim insertSQL As String
    Dim maxRowCount As Integer
    Dim DB�ڑ����, ����ݒ��� As String
    
    maxRowCount = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Count
    For i = 1 To maxRowCount
        If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON Then
            DB�ڑ���� = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
        End If
    Next
    
    maxRowCount = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).UsedRange.Count
    For i = 1 To maxRowCount
        ����ݒ��� = ����ݒ��� & addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(i, 3) & vbCrLf
    Next
    

    Set ADOConnection = connAccessDB()
    
    insertSQL = "INSERT INTO [Ddata�G���[�������] ([DB�ڑ����],[����ݒ���],[�G���[���],[�G���[�����[��],[�����N����],[�폜�t���O]) VALUES ('" & DB�ڑ���� & "','" & ����ݒ��� & "','" & errorMsg & "', '" & Environ("COMPUTERNAME") & "' ,'" & Format(Now, "yyyymmdd") & "','0')"
    
    ADOConnection.Execute (insertSQL)
    
    ADOConnection.Close
errorHandler:
    '�Ȃ������Ȃ�
    If Err.Number <> 0 Then
        errorMsg = errorMsg & "�G���[�ԍ��F" & Err.Number & vbNewLine
        errorMsg = errorMsg & "�G���[���e�F" & Err.Description & vbNewLine
        errorMsg = errorMsg & "�w���v�t�@�C�����F" & Err.HelpContext & vbNewLine
        errorMsg = errorMsg & "�v���W�F�N�g���F" & Err.Source & vbNewLine
        MsgBox errorMsg
    End If
    
End Function
