Attribute VB_Name = "CreateSQL"
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : カラム定義型より、SQLの値の型を変更する
' 引数   : String() カラムの定義情報  String カラムの値
' 戻り値 : 型変更後の値
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
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、InsertSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの論理名、String テーブルの物理名、一次元配列 カラム物理名、二次元配列 自動生成したデータ
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateInsertSqlSimple(tableNameJP As Variant, tableNameEN As Variant, columnNameEnArray As Variant, dataSetArray As Variant)
    'コメントを出力する。
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    'カラム作成
    Dim insertSqlFront As String
    insertSqlFront = "INSERT INTO " & UCase(tableNameEN) & "(" & Join(WorksheetFunction.Transpose(columnNameEnArray), ", ") & ") VALUES ("
    
    'バリュー作成
    For Each dataArray In dataSetArray
        insertSQL = insertSqlFront & "'" & Join(WorksheetFunction.Transpose(dataArray), "', '") & "');"
        PUB_TEMP_VAR_OBJ.WriteText insertSQL, adWriteLine
    Next
    
End Function
'***********************************************************************************************************************
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、UpdateSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの論理名、String テーブルの物理名、一次元配列 カラム物理名、二次元配列 自動生成したデータ、一次元配列 Where条件配列
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateUpdateSqlSimple(tableNameJP As Variant, tableNameEN As Variant, columnNameEnArray As Variant, dataSetArray As Variant, whereSqlSetArray As Variant)
    'コメントを出力する。
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    'カラム作成
    Dim updateSqlFront As String
    updateSqlFront = "UPDATE " & UCase(tableNameEN) & " SET "
    
    Dim updateColumnName As String
    Dim updateColumnData As Variant
    Dim updateSqlSet As Variant
    Dim updateSqlSetArray() As Variant
    
    Dim updateColumnDataRange As Range
    
    'セット配列に追加
    For columnNo = 1 To UBound(columnNameEnArray)
        updateColumnName = WorksheetFunction.Transpose(columnNameEnArray)(columnNo)
        Set updateColumnDataRange = Range(ActiveSheet.UsedRange.Cells(7, columnNo), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, columnNo))
        If updateColumnDataRange.Rows.Count > 1 Then
            updateColumnData = updateColumnDataRange
        Else
            ReDim updateColumnData(0)
            updateColumnData(0) = updateColumnDataRange.value
        End If
        
        'カラム名付け
        updateSqlSet = Split("', " & updateColumnName & " = '" & Join(WorksheetFunction.Transpose(updateColumnData), vbCrLf & "', " & updateColumnName & " = '"), vbCrLf)
        
        'サイズ定義
        ReDim Preserve updateSqlSetArray(UBound(updateSqlSet))
        
        For rowNo = 0 To UBound(updateSqlSet)
            updateSqlSetArray(rowNo) = updateSqlSetArray(rowNo) & updateSqlSet(rowNo)
        Next
    Next
    
    '出力
    For i = 0 To UBound(updateSqlSetArray)
        updateSQL = updateSqlFront & updateSqlSetArray(i) & "' WHERE" & whereSqlSetArray(i) & "';"
        updateSQL = Replace(updateSQL, "SET ',", "SET")
        updateSQL = Replace(updateSQL, "WHERE AND", "WHERE")
        PUB_TEMP_VAR_OBJ.WriteText updateSQL, adWriteLine
    Next
    
End Function
'***********************************************************************************************************************
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、DeleteSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの論理名、String テーブルの物理名、一次元配列 Where条件配列
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateDeleteSqlSimple(tableNameJP As Variant, tableNameEN As Variant, whereSqlSetArray As Variant)
    'コメントを出力する。
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/* " & tableNameJP & " " & tableNameEN & " */", adWriteLine

    'カラム作成
    Dim deleteSqlFront As String
    deleteSqlFront = "DELETE FROM " & UCase(tableNameEN) & " WHERE"
    
    'deleteSQL作成
    For Each whereSql In whereSqlSetArray
        deleteSQL = deleteSqlFront & whereSql & "';"
        deleteSQL = Replace(deleteSQL, "WHERE AND", "WHERE")
        PUB_TEMP_VAR_OBJ.WriteText deleteSQL, adWriteLine
    Next
    
End Function

Function CreateWhereSql() As Variant
    
    Dim findRange As Range
    Dim firstRange As Range
    Set findRange = ActiveSheet.UsedRange.Find("○")
    
    If Not findRange Is Nothing Then
    
        Dim whereColumnName As String
        Dim whereColumnData As Variant
        Dim whereSqlSet As Variant
        Dim whereSqlSetArray As Variant
        
        Dim whereColumnDataRange As Range
        
        'セル色が設定される場合
        If ActiveSheet.UsedRange.Cells(2, findRange.Column).Interior.Color = RGB(279, 117, 14) Then
            '見つかった１個目セルを処理する
            whereColumnName = ActiveSheet.UsedRange.Cells(3, findRange.Column).value
            Set whereColumnDataRange = ActiveSheet.Range(ActiveSheet.UsedRange.Cells(7, findRange.Column), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, findRange.Column))
            If whereColumnDataRange.Rows.Count > 1 Then
                whereColumnData = whereColumnDataRange
            Else
                ReDim whereColumnData(0, 0)
                whereColumnData(0, 0) = whereColumnDataRange.value
            End If
            
            'カラム名付け
            whereSqlSetArray = Split(" AND " & whereColumnName & " = '" & Join(WorksheetFunction.Transpose(whereColumnData), vbCrLf & " AND " & whereColumnName & " = '"), vbCrLf)
        End If

        
        Set firstRange = findRange
        Do
            Set findRange = ActiveSheet.UsedRange.FindNext(findRange)
            If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                Exit Do
            End If
            
            'セル色が設定される場合
            If ActiveSheet.UsedRange.Cells(2, findRange.Column).Interior.Color = RGB(279, 117, 14) Then
                '見つかったセルを処理する
                whereColumnName = ActiveSheet.UsedRange.Cells(3, findRange.Column).value
                Set whereColumnDataRange = Range(ActiveSheet.UsedRange.Cells(7, findRange.Column), ActiveSheet.UsedRange.Cells(ActiveSheet.UsedRange.Rows.Count, findRange.Column))
                If whereColumnDataRange.Rows.Count > 1 Then
                    whereColumnData = whereColumnDataRange
                Else
                    ReDim whereColumnData(0, 0)
                    whereColumnData(0, 0) = whereColumnDataRange.value
                End If
                
                'カラム名付け
                whereSqlSet = Split("' AND " & whereColumnName & " = '" & Join(WorksheetFunction.Transpose(whereColumnData), vbCrLf & "' AND " & whereColumnName & " = '"), vbCrLf)
                
                'セット配列に追加
                For i = 0 To UBound(whereSqlSet)
                    whereSqlSetArray(i) = whereSqlSetArray(i) & whereSqlSet(i)
                Next
            End If
            
        Loop While firstRange.Address <> findRange.Address
        
    End If
    
    CreateWhereSql = whereSqlSetArray
    
End Function

'***********************************************************************************************************************
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、InsertSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの物理名、String カラム物理名の配列、Collection データの配列の集合
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateInsertSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
      
    'コメントを出力する。
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/*" & PUB_TEMP_VAR_STR & " " & tableNameEN & "*/", adWriteLine
    
    'カラム作成
    Dim insertSqlFront As String
    insertSqlFront = "INSERT INTO " & UCase(tableNameEN) & "("
    
    Dim columns As String
    Dim columnName As Variant
    For Each columnName In tableColumns
        columns = columns & UCase(columnName) & ","
    Next
    
    '最後のコマを削除する
    columns = Left(columns, Len(columns) - 1)
    
    insertSqlFront = insertSqlFront & columns & ")"
    
    
    '特殊の項目値を修正するために、テーブル定義情報を取得する
    'カラム定義情報を取得する
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
   
    'バリュー作成
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
            
            'カラム型判断
            ReDim columnInfo(5)
            
            '特殊の項目値を修正する
            columnInfo = tabColumns(columnIndex)
            value = MakeSqlValue(columnInfo, CStr(data))
            
            Values = Values & value
            Values = Values & ","
            
            columnIndex = columnIndex + 1
        Next
        
        '最後のコマを削除する
        Values = Left(Values, Len(Values) - 1)
        insertSqlBehind = insertSqlBehind & Values & ");"
        insertSQL = insertSqlFront & insertSqlBehind

        PUB_TEMP_VAR_OBJ.WriteText insertSQL, adWriteLine
    Next
    

End Function

'***********************************************************************************************************************
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、DeleteSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの物理名、String カラム物理名の配列、Collection データの配列の集合
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateDeleteSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
    'コメントを出力する。
    PUB_TEMP_VAR_OBJ.WriteText vbCrLf & "/*" & PUB_TEMP_VAR_STR & " " & tableNameEN & "*/", adWriteLine
    
    'カラム作成
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
           
    'カラム定義情報を取得する
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
    
    'バリュー作成
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
            
            'deleteキー判断判断
            If deleteKeycolumns(i) <> "" Then
                deleteSqlBehind = deleteSqlBehind & deleteKeycolumns(i) & " = " & value & " AND "
            End If
    
        Next
        
        If deleteSqlBehind <> "" Then
            '最後の" AND "を削除する
            deleteSqlBehind = Left(deleteSqlBehind, Len(deleteSqlBehind) - 5)
            deleteSqlBehind = deleteSqlBehind & ";"
            deleteSQL = deleteSqlFront & deleteSqlBehind
            PUB_TEMP_VAR_OBJ.WriteText deleteSQL, adWriteLine
        End If
        
    Next
   
End Function

'***********************************************************************************************************************
' 機能   : Sql作成機能
' 概要   : シート内のデータを解析して、UpdateSQLを作成して、Sqlファイルに出力する
' 引数   : String テーブルの物理名、String カラム物理名の配列、Collection データの配列の集合
' 戻り値 : 無
'***********************************************************************************************************************
Public Function CreateUpdateSql(tableNameEN As String, tableColumns() As String, dataCollection As Collection)
    'コメントを出力する。
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
          
    
    'カラム定義情報を取得する
    Dim tabColumns As New Collection
    Dim columnInfo() As String
    Set tabColumns = GetTableColumnsNameFromDMRepository(tableNameEN)
    
    'バリュー作成
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
            
            '更新部を作成する
            updateSqlFront = updateSqlFront & tableColumns(i) & " =" & value & ", "
            
            '条件部を作成する
            If updateKeycolumns(i) <> "" Then
                updateSqlBehind = updateSqlBehind & updateKeycolumns(i) & " = " & value & " AND "
            End If
        Next
        
        If updateSqlBehind <> "" Then
            
            '最後の", "を削除する
            updateSqlFront = Left(updateSqlFront, Len(updateSqlFront) - 2)
            '最後の" AND "を削除する
            updateSqlBehind = Left(updateSqlBehind, Len(updateSqlBehind) - 5)
            updateSqlBehind = updateSqlBehind & ";"
            updateSQL = updateSqlFront & updateSqlBehind
            PUB_TEMP_VAR_OBJ.WriteText updateSQL, adWriteLine
        End If
        
    Next
   
End Function
