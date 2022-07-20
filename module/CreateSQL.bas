Attribute VB_Name = "CreateSQL"
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : �J������`�^���ASQL�̒l�̌^��ύX����
' ����   : String() �J�����̒�`���  String �J�����̒l
' �߂�l : �^�ύX��̒l
'***********************************************************************************************************************
Public Function MakeSqlValue(columnInfo() As String, value As String) As String
    Dim columnValue As String
    columnValue = ""
    
    If columnInfo(2) = "NUMBER" Then
        columnValue = value
    ElseIf columnInfo(2) = "DATE" Then
        If "SYSTIMESTAMP" = UCase(value) Or "SYSDATE" = UCase(value) Then
            Values = Values & "SYSDATE"
        Else
            If Len(value) = 14 Then
                columnValue = "TO_DATE('" & value & "','YYYY-MM-DD HH24MISS')"
            ElseIf Len(value) = 12 Then
                columnValue = "TO_DATE('" & value & "','YYYY-MM-DD HH24MI')"
            ElseIf Len(value) = 10 Then
                columnValue = "TO_DATE('" & value & "','YYYY-MM-DD HH')"
            ElseIf Len(value) = 8 Then
                columnValue = "TO_DATE('" & value & "','YYYY-MM-DD')"
            ElseIf Len(value) > 14 Then
                columnValue = "TO_DATE('" & value & "','YYYY-MM-DD HH24MISS')"
            ElseIf Len(value) = 0 Then
                columnValue = "NULL"
            ElseIf Len(value) < 8 Then
                columnValue = "SYSDATE"
            End If
        End If
    
    
    ElseIf columnInfo(2) = "TIMESTAMP" Then
        If "SYSTIMESTAMP" = UCase(value) Or "SYSDATE" = UCase(value) Then
            columnValue = "SYSTIMESTAMP"
        ElseIf IsEmpty(value) = True Then
            columnValue = "NULL"
        Else
            If Len(value) = 14 Then
                columnValue = "TO_TIMESTAMP('" & value & "','YYYY-MM-DD HH24MISS')"
            ElseIf Len(value) = 12 Then
                columnValue = "TO_TIMESTAMP('" & value & "','YYYY-MM-DD HH24MI')"
            ElseIf Len(value) = 10 Then
                columnValue = "TO_TIMESTAMP('" & value & "','YYYY-MM-DD HH')"
            ElseIf Len(value) = 8 Then
                columnValue = "TO_TIMESTAMP('" & value & "','YYYY-MM-DD')"
            ElseIf Len(value) > 14 Then
                columnValue = "TO_TIMESTAMP('" & value & "','YYYY-MM-DD HH24MISS')"
            ElseIf Len(value) = 0 Then
                columnValue = "NULL"
            ElseIf Len(value) < 8 Then
                columnValue = "SYSTIMESTAMP"
            End If
        End If
    ElseIf columnInfo(2) = "CLOB" Then
        columnValue = "TO_CLOB('" & value & "')"
    Else
        If "SYSTIMESTAMP" = UCase(value) Or "SYSDATE" = UCase(value) Then
            If columnInfo(3) = "14" Then
                columnValue = "TO_CHAR(" & value & ",'YYYYMMDDHH24MISS')"
            ElseIf columnInfo(3) = "12" Then
                columnValue = "TO_CHAR(" & value & ",'YYYYMMDDHH24MI')"
            ElseIf columnInfo(3) = "10" Then
                columnValue = "TO_CHAR(" & value & ",'YYYYMMDDHH')"
            ElseIf columnInfo(3) = "8" Then
                columnValue = "TO_CHAR(" & value & ",'YYYYMMDD')"
            ElseIf columnInfo(3) = "6" Then
                columnValue = "TO_CHAR(" & value & ",'YYYYMM')"
            Else
                columnValue = "'" & value & "'"
            End If
        ElseIf IsEmpty(value) = True Then
            columnValue = "NULL"
        Else
            columnValue = "'" & value & "'"
        End If
    End If

    MakeSqlValue = columnValue
End Function

'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āAInsertSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̘_�����AString �e�[�u���̕������A�ꎟ���z�� �J�����������A�񎟌��z�� �������������f�[�^
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateInsertSqlSimple(tableNameJP As Variant, tableNameEN As Variant, columnNameEnArray As Variant, dataSetArray As Variant)
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    '�J�����쐬
    Dim insertSqlFront As String
    insertSqlFront = "INSERT INTO " & UCase(tableNameEN) & "(" & Join(WorksheetFunction.Transpose(columnNameEnArray), ", ") & ") VALUES ("
    
    '�o�����[�쐬
    For Each dataArray In dataSetArray
        insertSQL = insertSqlFront & "'" & Join(WorksheetFunction.Transpose(dataArray), "', '") & "');"
        PUB_TEMP_VAR_OBJ.WriteText insertSQL, adWriteLine
    Next
    
End Function
'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āAUpdateSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̘_�����AString �e�[�u���̕������A�ꎟ���z�� �J�����������A�񎟌��z�� �������������f�[�^�A�ꎟ���z�� Where�����z��
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateUpdateSqlSimple(tableNameJP As Variant, tableNameEN As Variant, columnNameEnArray As Variant, dataSetArray As Variant, whereSqlSetArray As Variant)
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    '�J�����쐬
    Dim updateSqlFront As String
    updateSqlFront = "UPDATE " & UCase(tableNameEN) & " SET "
    
    Dim updateColumnName As String
    Dim updateColumnData As Variant
    Dim updateSqlSet As Variant
    Dim updateSqlSetArray() As Variant
    
    Dim updateColumnDataRange As Range
    
    '�Z�b�g�z��ɒǉ�
    For columnNo = 1 To UBound(columnNameEnArray)
        updateColumnName = WorksheetFunction.Transpose(columnNameEnArray)(columnNo)
        Set updateColumnDataRange = Range(ActiveSheet.UsedRange.Cells(7, columnNo), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, columnNo))
        If updateColumnDataRange.Rows.Count > 1 Then
            updateColumnData = updateColumnDataRange
        Else
            ReDim updateColumnData(0)
            updateColumnData(0) = updateColumnDataRange.value
        End If
        
        '�J�������t��
        updateSqlSet = Split("', " & updateColumnName & " = '" & Join(WorksheetFunction.Transpose(updateColumnData), vbCrLf & "', " & updateColumnName & " = '"), vbCrLf)
        
        '�T�C�Y��`
        ReDim Preserve updateSqlSetArray(UBound(updateSqlSet))
        
        For rowNo = 0 To UBound(updateSqlSet)
            updateSqlSetArray(rowNo) = updateSqlSetArray(rowNo) & updateSqlSet(rowNo)
        Next
    Next
    
    '�o��
    For i = 0 To UBound(updateSqlSetArray)
        updateSQL = updateSqlFront & updateSqlSetArray(i) & "' WHERE" & whereSqlSetArray(i) & "';"
        updateSQL = Replace(updateSQL, "SET ',", "SET")
        updateSQL = Replace(updateSQL, "WHERE AND", "WHERE")
        PUB_TEMP_VAR_OBJ.WriteText updateSQL, adWriteLine
    Next
    
End Function
'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āADeleteSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̘_�����AString �e�[�u���̕������A�ꎟ���z�� Where�����z��
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateDeleteSqlSimple(tableNameJP As Variant, tableNameEN As Variant, whereSqlSetArray As Variant)
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    '�J�����쐬
    Dim deleteSqlFront As String
    deleteSqlFront = "DELETE FROM " & UCase(tableNameEN) & " WHERE"
    
    'deleteSQL�쐬
    For Each whereSql In whereSqlSetArray
        deleteSQL = deleteSqlFront & whereSql & "';"
        deleteSQL = Replace(deleteSQL, "WHERE AND", "WHERE")
        PUB_TEMP_VAR_OBJ.WriteText deleteSQL, adWriteLine
    Next
    
End Function

Function CreateWhereSql() As Variant
    
    Dim findRange As Range
    Dim firstRange As Range
    Set findRange = ActiveSheet.UsedRange.Find("��")
    
    If Not findRange Is Nothing Then
    
        Dim whereColumnName As String
        Dim whereColumnData As Variant
        Dim whereSqlSet As Variant
        Dim whereSqlSetArray As Variant
        
        Dim whereColumnDataRange As Range
        
        '�Z���F���ݒ肳���ꍇ
        If ActiveSheet.UsedRange.Cells(2, findRange.Column).Interior.Color = RGB(279, 117, 14) Then
            '���������P�ڃZ������������
            whereColumnName = ActiveSheet.UsedRange.Cells(3, findRange.Column).value
            Set whereColumnDataRange = ActiveSheet.Range(ActiveSheet.UsedRange.Cells(7, findRange.Column), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, findRange.Column))
            If whereColumnDataRange.Rows.Count > 1 Then
                whereColumnData = whereColumnDataRange
            Else
                ReDim whereColumnData(0, 0)
                whereColumnData(0, 0) = whereColumnDataRange.value
            End If
            
            '�J�������t��
            whereSqlSetArray = Split(" AND " & whereColumnName & " = '" & Join(WorksheetFunction.Transpose(whereColumnData), vbCrLf & " AND " & whereColumnName & " = '"), vbCrLf)
        End If

        
        Set firstRange = findRange
        Do
            Set findRange = ActiveSheet.UsedRange.FindNext(findRange)
            If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                Exit Do
            End If
            
            '�Z���F���ݒ肳���ꍇ
            If ActiveSheet.UsedRange.Cells(2, findRange.Column).Interior.Color = RGB(279, 117, 14) Then
                '���������Z������������
                whereColumnName = ActiveSheet.UsedRange.Cells(3, findRange.Column).value
                Set whereColumnDataRange = Range(ActiveSheet.UsedRange.Cells(7, findRange.Column), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, findRange.Column))
                If whereColumnDataRange.Rows.Count > 1 Then
                    whereColumnData = whereColumnDataRange
                Else
                    ReDim whereColumnData(0, 0)
                    whereColumnData(0, 0) = whereColumnDataRange.value
                End If
                
                '�J�������t��
                whereSqlSet = Split("' AND " & whereColumnName & " = '" & Join(WorksheetFunction.Transpose(whereColumnData), vbCrLf & "' AND " & whereColumnName & " = '"), vbCrLf)
                
                '�Z�b�g�z��ɒǉ�
                For i = 0 To UBound(whereSqlSet)
                    whereSqlSetArray(i) = whereSqlSetArray(i) & whereSqlSet(i)
                Next
            End If
            
        Loop While firstRange.Address <> findRange.Address
        
    End If
    
    CreateWhereSql = whereSqlSetArray
    
End Function

'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āAInsertSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̕������AString �J�����������̔z��ACollection �f�[�^�̔z��̏W��
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateInsertSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
      
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/*" & PUB_TEMP_VAR_STR & " " & tableNameEN & "*/", adWriteLine
    
    '�J�����쐬
    Dim insertSqlFront As String
    insertSqlFront = "INSERT INTO " & UCase(tableNameEN) & "("
    
    Dim columns As String
    Dim columnName As Variant
    For Each columnName In tableColumns
        columns = columns & UCase(columnName) & ","
    Next
    
    '�Ō�̃R�}���폜����
    columns = Left(columns, Len(columns) - 1)
    
    insertSqlFront = insertSqlFront & columns & ")"
    
    
    '����̍��ڒl���C�����邽�߂ɁA�e�[�u����`�����擾����
    '�J������`�����擾����
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
   
    '�o�����[�쐬
    Dim record As Variant
    'ReDim record(UBound(tableColumns))
    For Each record In dataCollection
        Dim insertSQL As String
        Dim insertSqlBehind As String
        Dim Values As String
        insertSQL = ""
        insertSqlBehind = " VALUES ("
        Values = ""
        Dim data As Variant
        
        Dim columnIndex As Integer
        columnIndex = 1
        
        For Each data In record
            
            '�J�����^���f
            ReDim columnInfo(5)
            
            '����̍��ڒl���C������
            columnInfo = tabColumns(columnIndex)
            value = MakeSqlValue(columnInfo, CStr(data))
            
            Values = Values & value
            Values = Values & ","
            
            columnIndex = columnIndex + 1
        Next
        
        '�Ō�̃R�}���폜����
        Values = Left(Values, Len(Values) - 1)
        insertSqlBehind = insertSqlBehind & Values & ");"
        insertSQL = insertSqlFront & insertSqlBehind

        PUB_TEMP_VAR_OBJ.WriteText insertSQL, adWriteLine
    Next
    

End Function

'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āADeleteSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̕������AString �J�����������̔z��ACollection �f�[�^�̔z��̏W��
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateDeleteSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/*" & PUB_TEMP_VAR_STR & " " & tableNameEN & "*/", adWriteLine
    
    '�J�����쐬
    Dim deleteSqlFront As String
    deleteSqlFront = "DELETE FROM " & tableNameEN & " WHERE "
    
    
    Dim tabColumnsCollection As New Collection
    Set tabColumnsCollection = GetTabColumns(tableNameEN)
    
    If tabColumnsCollection Is Nothing Then
        Exit Function
    End If
    
    
    Dim deleteKeycolumns() As String
    Dim columnName As Variant
    Dim columnIndex As Integer
    columnIndex = 0
    ReDim deleteKeycolumns(UBound(tableColumns))
    For Each columnName In tableColumns
        Dim isKey As Boolean
        isKey = False
        For i = 1 To tabColumnsCollection.Count
            Dim tabColumnsArray() As String
            ReDim tabColumnsArray(1)
            tabColumnsArray = tabColumnsCollection(i)
            
            If tabColumnsArray(0) = columnName And tabColumnsArray(1) = "Y" Then
                isKey = True
                Exit For
            End If
        Next
        
        If isKey = True Then
                
            deleteKeycolumns(columnIndex) = columnName
        
        End If
        
        columnIndex = columnIndex + 1
    Next
           
    '�J������`�����擾����
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
    
    '�o�����[�쐬
    Dim record As Variant
    i = 0
    ReDim record(UBound(tableColumns))
    For Each record In dataCollection
        Dim deleteSqlBehind As String
        Dim deleteSQL As String
        deleteSqlBehind = ""
        deleteSQL = ""
        For i = 0 To UBound(tableColumns)
            
            
            ReDim columnInfo(5)
            columnInfo = tabColumns(i + 1)
            value = MakeSqlValue(columnInfo, CStr(record(i)))
            
            'delete�L�[���f���f
            If deleteKeycolumns(i) <> "" Then
                deleteSqlBehind = deleteSqlBehind & deleteKeycolumns(i) & " = " & value & " AND "
            End If
    
        Next
        
        If deleteSqlBehind <> "" Then
            '�Ō��" AND "���폜����
            deleteSqlBehind = Left(deleteSqlBehind, Len(deleteSqlBehind) - 5)
            deleteSqlBehind = deleteSqlBehind & ";"
            deleteSQL = deleteSqlFront & deleteSqlBehind
            PUB_TEMP_VAR_OBJ.WriteText deleteSQL, adWriteLine
        End If
        
    Next
   
End Function

'***********************************************************************************************************************
' �@�\   : Sql�쐬�@�\
' �T�v   : �V�[�g���̃f�[�^����͂��āAUpdateSQL���쐬���āASql�t�@�C���ɏo�͂���
' ����   : String �e�[�u���̕������AString �J�����������̔z��ACollection �f�[�^�̔z��̏W��
' �߂�l : ��
'***********************************************************************************************************************
Public Function CreateUpdateSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
    '�R�����g���o�͂���B
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/*" & PUB_TEMP_VAR_STR & " " & tableNameEN & "*/", adWriteLine
    
    Dim tabColumnsCollection As New Collection
    Set tabColumnsCollection = GetTabColumns(tableNameEN)
    
    If tabColumnsCollection Is Nothing Then
        Exit Function
    End If
    
    Dim updateKeycolumns() As String
    Dim columnName As Variant
    Dim columnIndex As Integer
    columnIndex = 0
    ReDim updateKeycolumns(UBound(tableColumns))
    For Each columnName In tableColumns
        Dim isKey As Boolean
        isKey = False
        For i = 1 To tabColumnsCollection.Count
            Dim tabColumnsArray() As String
            ReDim tabColumnsArray(1)
            tabColumnsArray = tabColumnsCollection(i)
            
            If tabColumnsArray(0) = columnName And tabColumnsArray(1) = "Y" Then
                isKey = True
                Exit For
            End If

        Next
        
        If isKey = True Then
                
            updateKeycolumns(columnIndex) = columnName
        
        End If
        
        columnIndex = columnIndex + 1
    Next
          
    
    '�J������`�����擾����
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
    '�o�����[�쐬
    Dim record As Variant
    i = 0
    ReDim record(UBound(tableColumns))
    For Each record In dataCollection
        Dim updateSqlBehind As String
        Dim deleteSQL As String
        Dim updateSqlFront As String
        updateSqlFront = "UPDATE " & tableNameEN & " SET "
        updateSqlBehind = " WHERE "
        updateSQL = ""
        For i = 0 To UBound(tableColumns)
            
            ReDim columnInfo(5)
            columnInfo = tabColumns(i + 1)
            value = MakeSqlValue(columnInfo, CStr(record(i)))
            
            '�X�V�����쐬����
            updateSqlFront = updateSqlFront & tableColumns(i) & " =" & value & ", "
            
            '���������쐬����
            If updateKeycolumns(i) <> "" Then
                updateSqlBehind = updateSqlBehind & updateKeycolumns(i) & " = " & value & " AND "
            End If
        Next
        
        If updateSqlBehind <> "" Then
            
            '�Ō��", "���폜����
            updateSqlFront = Left(updateSqlFront, Len(updateSqlFront) - 2)
            '�Ō��" AND "���폜����
            updateSqlBehind = Left(updateSqlBehind, Len(updateSqlBehind) - 5)
            updateSqlBehind = updateSqlBehind & ";"
            updateSQL = updateSqlFront & updateSqlBehind
            PUB_TEMP_VAR_OBJ.WriteText updateSQL, adWriteLine
        End If
        
    Next
   
End Function
