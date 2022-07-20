Attribute VB_Name = "CreateData"
'***********************************************************************************************************************
' 機能   : 試験データ作成機能
' 概要   : 論理名から物理名に変換する
' 引数   : String　論理名
' 戻り値 : String　物理名
'***********************************************************************************************************************
Public Function CreateTestData()

    On Error GoTo errorHandler
    
    'D-Tools画面の初期設定を実施する
    Load DTools
    
    '入力情報を保存する
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3) = ActiveSheet.Cells(1, 2).value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(19, 3) = ActiveSheet.Cells(2, 2).value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(20, 3) = ActiveSheet.Cells(3, 2).value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(21, 3) = ActiveSheet.Cells(4, 2).value
    
    '作成対象テーブル
    Dim tableNames As String
    tableNames = ActiveSheet.Cells(1, 2).value
    'レコード数
    Dim totleRecordCount As Integer
    totleRecordCount = ActiveSheet.Cells(2, 2).value
    '番号の枝番
    Dim countFromNo As String
    countFromNo = ActiveSheet.Cells(4, 2).value
    
    Dim usedRowCount As Integer
    usedRowCount = ActiveSheet.UsedRange.Rows.Count

    
    Dim columnNameJP As String
    Dim columnValue As String
    Dim dataRowStartIndex As Integer
    dataRowStartIndex = 6
    Do While dataRowStartIndex <= usedRowCount
        If dataRowStartIndex = usedRowCount Then
            columnNameJP = columnNameJP & ActiveSheet.Cells(dataRowStartIndex, 1)
        Else
            columnNameJP = columnNameJP & ActiveSheet.Cells(dataRowStartIndex, 1) & ","
        End If
        dataRowStartIndex = dataRowStartIndex + 1
    Loop
    dataRowStartIndex = 6
    Do While dataRowStartIndex <= usedRowCount
        If dataRowStartIndex = usedRowCount Then
            columnValue = columnValue & ActiveSheet.Cells(dataRowStartIndex, 2)
        Else
            columnValue = columnValue & ActiveSheet.Cells(dataRowStartIndex, 2) & ","
        End If
        dataRowStartIndex = dataRowStartIndex + 1
    Loop
    
    '入力情報を保存する
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(22, 3) = columnNameJP
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(23, 3) = columnValue
    
    If DATA_SOURCE_DIR = "" Then
        MsgBox "DM定義リポジトリを設定してください。"
        Exit Function
    End If
    
    '作業用シートを作成する
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, ActiveWorkbook)
       
    '試験データを作成する
    Dim tableName As Variant
    Dim rowIndex As Integer
    rowIndex = 1
    For Each tableName In Split(tableNames, ",")
        'テーブル名を取得する
        Dim tableNameInfoCollection As New Collection
        Set tableNameInfoCollection = GetTableNameFromDMRepository(CStr(tableName))
        
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
                '罫線を付ける
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address, Cells(rowIndex, 2).Address).Borders.LineStyle = xlContinuous
                
                'テーブルカラム情報を取得する
                Set columnsInfoCollection = GetTableColumnsNameFromDMRepository(tableNameInfo(0))
                
                colIndex = 1
                Do While colIndex <= columnsInfoCollection.Count
                    ReDim columnInfo(5)
                    
                    columnInfo = columnsInfoCollection(colIndex)
                    
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 1, colIndex).value = columnInfo(0)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 2, colIndex).value = columnInfo(1)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 3, colIndex).value = columnInfo(2)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 4, colIndex).value = columnInfo(3)
                    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 5, colIndex).value = columnInfo(4)
                    If columnInfo(5) <> "" Then
                        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 1, colIndex).Interior.Color = RGB(279, 117, 14)
                    End If
                    
                    Dim countNowNo As String
                    countNowNo = countFromNo
                    
                    For i = 1 To totleRecordCount
                        Dim value As String
                        value = CreateColumnDataByUserDefine(countNowNo, columnInfo)
                        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(rowIndex + 5 + i, colIndex).value = value
                        countNowNo = NumberCountUpToStr(CInt(countNowNo), 1, 3)
                    Next
                                    
                    colIndex = colIndex + 1
                Loop
                
                '罫線を付ける
                 OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex + 1, 1).Address & ":" & Cells(rowIndex + 5 + totleRecordCount, colIndex - 1).Address).Borders.LineStyle = xlContinuous
                '列の幅自動調整
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex + 2, colIndex - 1).Address).columns.AutoFit
        
                rowIndex = rowIndex + 8 + totleRecordCount
                
                tableNameIndex = tableNameIndex + 1
                
            Loop
        
        End If
        
    Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("CreateTestData")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If
    
End Function

'***********************************************************************************************************************
' 機能   : 試験データ作成機能
' 概要   : 設定の情報より、カラムの値を作成する。
' 引数   : String 個人識別子, String 枝番, String 指定論理カラム名, String 指定カラム値, String カラム定義情報
' 戻り値 : String　カラム値
'***********************************************************************************************************************
Public Function CreateColumnDataByUserDefine(countNo, columnInfo() As String) As String
    
    Dim tableNames As String
    Dim totleRecordCount As String
    Dim memberKey As String

    Dim columnNameJP As String
    Dim columnValue As String
    Dim columnNameJPArr As Variant
    Dim columnValueArr As Variant
    
    Dim value As String
    Dim columnName As Variant
    
    
    tableNames = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3)
    totleRecordCount = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(19, 3)
    memberKey = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(20, 3)
    columnNameJP = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(22, 3)
    columnValue = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(23, 3)
    
    columnNameJPArr = Split(columnNameJP, ",")
    columnValueArr = Split(columnValue, ",")
    
    '指定値を設定する
    For i = 0 To UBound(columnNameJPArr)
        If columnInfo(0) Like columnNameJPArr(i) Then
            value = columnValueArr(i)
            If Len(value) >= 7 And Not (columnInfo(0) Like "*[タイムスタンプ,時分秒,年月日,年月,日付]") Then
                value = Mid(value, 1, Len(value) - 3) & countNo
            End If
            Exit For
        End If
    Next
    
    '指定値以外の場合
    If value = "" Then
        If columnInfo(2) Like "*[CHAR,VARCHAR2,NUMBER,CLOB]" Then
            
            Dim columnValueKey As String
            columnValueKey = GetColumnValueKey(columnInfo(0))
            
            Dim valueLen As String
            If columnInfo(3) > 20 Then
                valueLen = 20
            Else
                valueLen = columnInfo(3)
            End If

            If valueLen < 7 Then
                 For j = 1 To valueLen
                    value = value & "0"
                Next
            Else
                value = columnValueKey & memberKey
                For j = 1 To valueLen - 7
                    value = value & "0"
                Next
                value = value & countNo
            End If
            
       
        ElseIf columnInfo(2) = "TIMESTAMP" Then
            value = "SYSTIMESTAMP"
        Else
            If columnInfo(4) = "NULL不可" Then
                value = " "
            Else
                value = ""
            End If
        End If
        
    End If
    
    CreateColumnDataByUserDefine = value
    
End Function

'***********************************************************************************************************************
' 機能   : 試験データ作成機能
' 概要   : カラム和名より、カラムの識別子を取得する。
' 引数   : String カラム和名
' 戻り値 : String　カラムの識別子
'***********************************************************************************************************************
Private Function GetColumnValueKey(columnNameJP As String) As String
    Dim columnValueKey As String
    
    Dim sql As String
    sql = "SELECT [識別子コード] FROM [識別子管理テーブル] WHERE [識別子名] = '" & columnNameJP & "' AND [識別子種別] = '1' AND [削除フラグ] = '0'"
       
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Set ADORecordset = New ADODB.recordset
    ADORecordset.Open sql, ADOConnection
    
    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("識別子コード")) Then
            columnValueKey = ""
        Else
            columnValueKey = resultFields("識別子コード").value
        End If
        
        ADORecordset.MoveNext
    Loop
    
    If columnValueKey = "" Then
        columnValueKey = "00"
    End If
    
    GetColumnValueKey = columnValueKey
End Function

'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 番号をカウントアップした後、固定長にする
' 引数   : Integer カラム和名
' 戻り値 : String　カラムの識別子
'***********************************************************************************************************************
Private Function NumberCountUpToStr(startNo As Integer, countUp As Integer, StrLen As Integer) As String
    Dim afterCountUpNo As Integer
    Dim afterCountUpStr As Integer
    
    afterCountUpNo = startNo + countUp
    afterCountUpStr = afterCountUpNo & ""
    For i = 0 To StrLen - Len(afterCountUpStr)
        afterCountUpStr = "0" & afterCountUpStr
    Next
    NumberCountUpToStr = afterCountUpStr
End Function
