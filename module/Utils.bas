Attribute VB_Name = "Utils"
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : アドインに操作データを格納するためのシートをを作成する。初期化時のみ。
' 引数   : String シート名、Workbook　追加対象のエクセル
' 戻り値 : 新シート名
'***********************************************************************************************************************
Public Function createNewSheet(resultSheetName As String, addWorkbook As Workbook) As String
    Dim existFlag As Boolean
    existFlag = checkSheetNameExist(resultSheetName, addWorkbook)
       
    '新シート作成
    If existFlag = False Then
        addWorkbook.Sheets.Add(After:=Worksheets(Worksheets.Count)).name = resultSheetName
        '表示形式を文字列に設定する
        addWorkbook.Worksheets(resultSheetName).Cells.NumberFormatLocal = "@"
    End If
    
    createNewSheet = addWorkbook.Worksheets(Worksheets.Count).name
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 既存のシート数、シート名より、新しいシートを追加する。
' 引数   : String シート名、Workbook　追加対象のエクセル
' 戻り値 : 無
'***********************************************************************************************************************
Public Function outPutResultSheet(resultSheetName As String, addWorkbook As Workbook)
    Dim existFlag As Boolean
    existFlag = False
    
    Dim sheetsCount As Integer
    sheetsCount = addWorkbook.Worksheets.Count
    
    existFlag = checkSheetNameExist(resultSheetName & sheetsCount + 1, addWorkbook)
    
    '新シート作成
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
    '表示形式を文字列に設定する
    addWorkbook.Worksheets(resultSheetName).Cells.NumberFormatLocal = "@"
    '出力結果のシート名
    RESULT_SHEET_NAME = resultSheetName
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 指定シート名を指定エクセルに存在することをチェックする
' 引数   : String シート名、Workbook　指定のエクセル
' 戻り値 : チェック結果（TRUE:存在／FALSE：不存在）
'***********************************************************************************************************************
Public Function checkSheetNameExist(resultSheetName As String, addWorkbook As Workbook) As Boolean
    Dim existFlag As Boolean
    existFlag = False
    'シート存在チェック
    For i = 1 To addWorkbook.Sheets.Count
        If resultSheetName = addWorkbook.Sheets(i).name Then
            existFlag = True
            Exit For
        End If
    Next
checkSheetNameExist = existFlag
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 全体変数から、接続情報を利用して、DB接続する
' 引数   : なし
' 戻り値 : DB接続のオブジェクト
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
' 機能   : 共通機能
' 概要   : 全体変数から、接続情報を利用して、AccessDB接続する
' 引数   : なし
' 戻り値 : DB接続のオブジェクト
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
' 機能   : 共通機能
' 概要   : DBからテーブルの物理名より、テーブルの物理、論理カラム名を取得する
' 引数   : String　テーブルの物理名
' 戻り値 : カラム物理名と論理名の結果集合
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
                    columnName(4) = "NULL不可"
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
' 機能   : 共通機能
' 概要   : DM定義リポジトリからテーブルの物理名より、テーブルの物理、論理カラム名を取得する
' 引数   : String　テーブルの物理名
' 戻り値 : カラム物理名と論理名の結果集合
'***********************************************************************************************************************
Public Function GetTableColumnsNameFromDMRepository(tableNameEN As String) As Collection
    Dim sql As String
    sql = "SELECT [属性名_和名],[カラム名_英名],[データ型],[桁数],[主キー],[NULL] FROM [属性定義書] WHERE [テーブル名_英名] = '" & UCase(tableNameEN) & "' ORDER BY [No]"
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Dim ADORecordset As New ADODB.recordset
    Set ADORecordset = ADOConnection.Execute(sql)
    
    Dim columnNameCollection As New Collection
    
    '0件チェック
    Do Until ADORecordset.EOF
        Dim columnName() As String
        ReDim columnName(5)
        Set resultFields = ADORecordset.Fields
        columnName(0) = resultFields("属性名_和名").value
        columnName(1) = resultFields("カラム名_英名").value
        columnName(2) = resultFields("データ型").value
        columnName(3) = resultFields("桁数").value
        columnName(4) = resultFields("NULL").value
        columnName(5) = resultFields("主キー").value
        columnNameCollection.Add (columnName)
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    '取得結果を返却する
    Set GetTableColumnsNameFromDMRepository = columnNameCollection
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : テーブル物理名の場合、DM定義リポジトリからテーブル論理名取得する。テーブル論理名の場合、DM定義リポジトリからテーブル物理名を取得する。
' 引数   : String　テーブル名
' 戻り値 : Collection テーブル論理名とテーブル物理名の集合
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
        sql = "SELECT [テーブル名_和名], [テーブル名_英名] FROM [テーブル定義書] WHERE [テーブル名_和名] LIKE '" & tableName & "'"
    Else
        tableName = Replace(tableName, "*", "%")
        sql = "SELECT [テーブル名_和名], [テーブル名_英名] FROM [テーブル定義書] WHERE [テーブル名_英名] LIKE '" & UCase(tableName) & "'"
    End If
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Dim ADORecordset As New ADODB.recordset
    Set ADORecordset = ADOConnection.Execute(sql)
    
    Do Until ADORecordset.EOF
        ReDim tableNameInfo(1)
        Set resultFields = ADORecordset.Fields
        tableNameEN = resultFields("テーブル名_英名").value
        tableNameJP = resultFields("テーブル名_和名").value
        
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
' 機能   : 共通機能
' 概要   : テーブル物理名の場合、DBからテーブル論理名取得する。テーブル論理名の場合、DBからテーブル物理名を取得する。
' 引数   : String　テーブル名
' 戻り値 : Collection テーブル論理名とテーブル物理名の集合
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
    If StrConv(Application.GetPhonetic(Replace(UCase(tableName), "ID", "")), vbHiragana) Like "*[あ-ん]*" Then
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
' 機能   : 共通機能
' 概要   : DBにあるのすべてのユーザーのテーブルの物理名を取得する
' 引数   : 無
' 戻り値 : テーブル物理名の集合
'***********************************************************************************************************************
Public Function GetAllTableNameEN() As Collection
    Dim sql As String
    Dim resutList As New Collection
    sql = "SELECT テーブル名_英名 AS TABLE_NAME FROM テーブル定義書"
    
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
    
    '取得結果を返却する
    Set GetAllTableNameEN = resutList
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 指定行に、テーブル物理名を存在していることをチェックする
' 引数   : Integer　行番号
' 戻り値 : チェック結果(TRUE:存在/FALSE：不存在)
'***********************************************************************************************************************
Public Function checkTableNameExistInRow(rowNo As Integer) As Boolean
    'テーブル物理名を探す
    Dim tableNameEN As String
    Dim rowRange As Range
    
    Dim allTableNameEN As New Collection
    Set allTableNameEN = GetAllTableNameEN
    
    Dim existFlag As Boolean
    
    For Each rowRange In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
        existFlag = False
        If StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[あ-ん]*" Then
            '日本語を含む場合、同行横の次のセルをチェックする。
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
' 機能   : 共通機能
' 概要   : 指定行に、指定文字（複数可）を存在していることをチェックする
' 引数   : Integer　行番号、String 指定の文字配列
' 戻り値 : チェック結果(0:不存在 1:存在あり 2：全部存在)
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
' 機能   : 共通機能
' 概要   : 指定テーブル物理名のカラム情報を取得する
' 引数   : String　テーブルの物理名
' 戻り値 : カラムごとのカラム物理名、NULL可否の配列の集合
'***********************************************************************************************************************
Public Function GetTabColumns(tableNameEN As String) As Collection
    Dim sql As String
    Dim resutList As New Collection
    sql = "SELECT [カラム名_英名] AS COLUMN_NAME,IIF([主キー] <> '','Y','N') AS ISKEY FROM [属性定義書] WHERE [テーブル名_英名] = '" & UCase(tableNameEN) & "'"
    
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
    
    '取得結果を返却する
    Set GetTabColumns = resutList

End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : DBから指定テーブル物理名のカラム情報を取得する
' 引数   : String　テーブルの物理名
' 戻り値 : カラムごとのカラム物理名、NULL可否、型、サイズの配列の集合
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
    
    '取得結果を返却する
    Set GetTabColumnInfo = resutList
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 文字列内の指定の文字列から、指定の文字列までの文字列を切り取る
' 引数   : String 切取対象の文字列、String 切取開始文字列、String 切取終了文字列
' 戻り値 : 切取文字列
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
' 機能   : 共通機能
' 概要   : 文字列内のspace、タブコード、改行コードを空白に置換する
' 引数   : 置換対象の文字列
' 戻り値 : 置換後の文字列
'***********************************************************************************************************************
Public Function ReplaceSTNToNull(str As String) As String
    '空白
    str = Replace(str, Chr(32), "")
    'タブ
    str = Replace(str, Chr(9), "")
    '改行LF
    str = Replace(str, Chr(10), "")
    '改行CR
    str = Replace(str, Chr(13), "")
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 文字列内のタブコード、改行コードを空白に置換する
' 引数   : 置換対象の文字列
' 戻り値 : 置換後の文字列
'***********************************************************************************************************************
Public Function ReplaceTNToSpace(str As String) As String
    'タブ
    str = Replace(str, Chr(9), " ")
    '改行LF
    str = Replace(str, Chr(10), " ")
    '改行CR
    str = Replace(str, Chr(13), " ")
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 指定行が空白の行かどうかをチェックする
' 引数   : Integer 行番号
' 戻り値 : チェック結果（TRUE：Allセルが空白／FALSE:空白以外のセルを存在する）
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
' 機能   : 共通機能
' 概要   : 指定行の一個目空白ではないセルを取得する
' 引数   : Integer 行番号
' 戻り値 : Range セルオブジェクト
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
' 機能   : 共通機能
' 概要   : 指定セル下の同列に、一個目空白ではないセルを取得する
' 引数   : Range セルオブジェクト
' 戻り値 : Range セルオブジェクト
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
' 機能   : 共通機能
' 概要   : 指定開始行から指定終了行まで、グループ化する
' 引数   : Integer 開始行　Integer　終了行
' 戻り値 : 無
'***********************************************************************************************************************
Function Group(groupStartIndex As Integer, groupEndIndex As Integer)
    On Error GoTo Continue
    ActiveSheet.Rows(groupStartIndex & ":" & groupEndIndex).Group
Continue:
End Function

'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 指定セルの内容より、グループ化ための開始点かどうかを判断する
' 引数   : String セルの内容
' 戻り値 : Boolean チェック結果
'***********************************************************************************************************************
Function IsGroupStartRow(cellValue As String) As Boolean
    Dim checkResult As Boolean
    checkResult = False
    
    If cellValue Like "■*" Then
        checkResult = True
    ElseIf cellValue Like "[(][0-9]*" Then
        checkResult = True
    ElseIf cellValue Like "[（][0-9]*" Then
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
' 機能   : Sql作成機能
' 概要   : 指定行からテーブル物理名を探して、返却する
' 引数   : Integer 行番号
' 戻り値 : 見つかったテーブル物理名
'***********************************************************************************************************************
Public Function CreateSql_GetTableName(rowNo As Integer) As String
    'テーブル物理名を探す
    Dim tableNameEN As String
    Dim rowRange As Range
    For Each rowRange In ActiveSheet.Rows(rowNo).SpecialCells(xlCellTypeConstants)
        If rowRange.value = "" Then
            '空白の場合、同行横の次のセルをチェックする。
            GoTo Continue
        ElseIf StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[あ-ん]*" Then
            '日本語を含む場合、同行横の次のセルをチェックする。
            'テーブル物理名と想定
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
' 機能   : Sql作成機能
' 概要   : 指定行がカラム行の場合、カラム名の集合、カラムの開始インデックス、カラムの終了インデックス
' 引数   : Integer 行番号
' 戻り値 : 返却したい結果の配列
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
                
                If StrConv(Application.GetPhonetic(Replace(UCase(rowRange.value), "ID", "")), vbHiragana) Like "*[あ-ん]*" Then
                    '日本語を含む場合、物理カラムではない
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
' 機能   : Sql作成機能
' 概要   : 指定行の指定列のデータを配列に実装して、返却する
' 引数   : Integer 行番号, columnStartIndex カラム開始インデックス, columnEndIndex カラム終了インデックス
' 戻り値 : 指定行の指定列のデータ配列
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
' 機能   : 共通機能
' 概要   : 引数が配列か判定し、配列の場合は空かどうかも判定する
' 引数   : varArray  配列
' 戻り値 : 判定結果（1:配列/0:空の配列/-1:配列じゃない）
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
' 機能   : 共通機能
' 概要   : 現在開いているファイルのディレクトリをチェックして、空白の場合、ディスクトップのディレクトリを返却する。
' 引数   : なし
' 戻り値 : ディレクトリの文字列
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
' 機能   : 共通機能
' 概要   : 識別子を取得する
' 引数   : String 識別子の種別（設定しない時、すべての識別を取得する）
' 戻り値 : collection　識別子コードと識別子名の集合
'***********************************************************************************************************************
Public Function GetKeyValue(keyCode As String) As Collection
    Dim keyValueInfo() As String
    Dim keyValueInfoCollection As New Collection
    Dim sql As String
    Dim keyName As String
    Dim keyValue As String
    
    If keyCode = "1" Then
        sql = "SELECT [識別子コード], [識別子名] FROM [識別子管理テーブル] WHERE [識別子種別] = '1' AND [削除フラグ] = '0'"
    ElseIf keyCode = "2" Then
        sql = "SELECT [識別子コード], [識別子名] FROM [識別子管理テーブル] WHERE [識別子種別] = '2' AND [削除フラグ] = '0'"
    ElseIf keyCode = "" Then
        sql = "SELECT [識別子コード], [識別子名] FROM [識別子管理テーブル] WHERE [削除フラグ] = '0'"
    Else
        Exit Function
    End If
    
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    Set ADORecordset = New ADODB.recordset
    ADORecordset.Open sql, ADOConnection
    
    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("識別子名")) Then
            keyName = ""
        Else
            keyName = resultFields("識別子名").value
        End If
        If IsNull(resultFields("識別子コード")) Then
            keyValue = ""
        Else
            keyValue = resultFields("識別子コード").value
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
' 機能   : バージョンチェック機能
' 概要   : 利用しているバージョンが最新かどうかチェックする
' 引数   : なし
' 戻り値 : なし
'***********************************************************************************************************************
Public Function CheckDdataVersion()
    On Error GoTo errorHandler

    'D-Tools画面の初期設定を実施する
    Load DTools
    
    Dim ADOConnection As New ADODB.Connection
    Dim dataSource As String
    dataSource = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DATA_SOURCE_DIR & ";"
    ADOConnection.Open dataSource
    
    '最新のバージョン情報を提示する
    Dim sql As String
    Dim ADORecordset As New ADODB.recordset
    Dim バージョン, 改修内容, 修正日, 修正者, アドイン格納場所 As String
    Dim rowCount As Integer
    Dim usedVersion As String
    Dim result As Integer
    
    rowCount = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).UsedRange.Rows.Count
    
    Do Until usedVersion <> ""
        usedVersion = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).Cells(rowCount, 2).value
        rowCount = rowCount - 1
    Loop
            
    sql = "SELECT [バージョン], [改修内容],[修正日],[修正者],[アドイン格納場所] FROM [Ddataバージョン情報] WHERE [ID] = (SELECT MAX([ID]) FROM [Ddataバージョン情報] WHERE [バージョン] > " & Val(usedVersion) & ")"
    ADORecordset.Open sql, ADOConnection

    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        
        If IsNull(resultFields("バージョン")) Then
            バージョン = ""
        Else
            バージョン = resultFields("バージョン").value
        End If
        
        If IsNull(resultFields("改修内容")) Then
            改修内容 = ""
        Else
            改修内容 = resultFields("改修内容").value
        End If
        
        If IsNull(resultFields("修正日")) Then
            修正日 = ""
        Else
            修正日 = resultFields("修正日").value
        End If
        
        If IsNull(resultFields("修正者")) Then
            修正者 = ""
        Else
            修正者 = resultFields("修正者").value
        End If
        
        If IsNull(resultFields("アドイン格納場所")) Then
            アドイン格納場所 = ""
        Else
            アドイン格納場所 = resultFields("アドイン格納場所").value
        End If
          
        Dim versionMessage  As String
        
        versionMessage = "D-Toolsの最新バージョン【" & バージョン & "】をリリースしました。" & vbCrLf & vbCrLf & _
                         "修正日：" & 修正日 & "  修正者：" & 修正者 & vbCrLf & vbCrLf & _
                         "アドイン格納場所：" & vbCrLf & アドイン格納場所 & vbCrLf & vbCrLf & _
                         "改修内容：" & vbCrLf & 改修内容
                         
        versionMessage = versionMessage & vbCrLf & vbCrLf & vbCrLf & "はい(Y)を押下すると、アドイン格納の場所をコピーします。"
        versionMessage = versionMessage & vbCrLf & "いいえ(N)を押下すると、今回の更新をスキップします。"
        
        result = MsgBox(versionMessage, vbYesNo + vbExclamation)
        
        If result = vbYes Then
            Dim myData As New DataObject
            myData.SetText (アドイン格納場所)
            myData.PutInClipboard
        ElseIf result = vbNo Then
            '改修履歴シートに、今の最新のバージョン情報を記載して、今後の更新提示をしなくなります。
            rowCount = addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).UsedRange.Rows.Count
            addInWorkBook.Worksheets(VERSION_HISTORY_SHEET_NAME).Cells(rowCount + 1, 2).value = バージョン
        End If
        
        ADORecordset.MoveNext
        
    Loop
    
errorHandler:
'D-Tools画面の初期設定を解放する
    Call CloseForm
    
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 対象文字列日本語が存在するかをチェック
' 引数   : チェック対象文字列
' 戻り値 : True：日本語文字が存在、False：日本語文字が不存在
'***********************************************************************************************************************
Public Function IsContainJapanese(str As String) As Boolean
    Dim charStr As String
    For i = 1 To Len(str)
        charStr = Mid(str, i, 1)
        If StrConv(Application.GetPhonetic(charStr), vbHiragana) Like "*[あ-ん]*" Then
            IsContainJapanese = True
            Exit Function
        End If
    Next
    
    IsContainJapanese = False
End Function


'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : エラーを発生した時、ダイアログにエラー内容を表示する。
' 引数   : なし
' 戻り値 : なし
'***********************************************************************************************************************
Public Function ShowErrorMsg(Optional functionName As String)
    Dim errorMsg As String
    
    errorMsg = errorMsg & "エラー番号：" & Err.Number & vbNewLine
    errorMsg = errorMsg & "エラー内容：" & Err.Description & vbNewLine
    errorMsg = errorMsg & "ヘルプファイル名：" & Err.HelpContext & vbNewLine
    errorMsg = errorMsg & "プロジェクト名：" & Err.Source & vbNewLine
    If functionName <> "" Then
        errorMsg = errorMsg & "メソッド名：" & functionName & vbNewLine
    End If
    Err.Clear
    
    '最初の2か月だけ、エラー情報を収集する
    If Format(Now, "yyyymmdd") < "20990120" Then
        Call SaveErrorInfo(errorMsg)
    End If
    
    MsgBox errorMsg
    
End Function

'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : エラーを発生した時、ダイアログにエラー内容を表示する。
' 引数   : なし
' 戻り値 : なし
'***********************************************************************************************************************
Public Function SaveErrorInfo(errorMsg As String)
    On Error GoTo errorHandler
    
    Dim ADOConnection As New ADODB.Connection
    Dim insertSQL As String
    Dim maxRowCount As Integer
    Dim DB接続情報, 操作設定情報 As String
    
    maxRowCount = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Count
    For i = 1 To maxRowCount
        If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON Then
            DB接続情報 = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
        End If
    Next
    
    maxRowCount = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).UsedRange.Count
    For i = 1 To maxRowCount
        操作設定情報 = 操作設定情報 & addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(i, 3) & vbCrLf
    Next
    

    Set ADOConnection = connAccessDB()
    
    insertSQL = "INSERT INTO [Ddataエラー発生情報] ([DB接続情報],[操作設定情報],[エラー情報],[エラー発生端末],[発生年月日],[削除フラグ]) VALUES ('" & DB接続情報 & "','" & 操作設定情報 & "','" & errorMsg & "', '" & Environ("COMPUTERNAME") & "' ,'" & Format(Now, "yyyymmdd") & "','0')"
    
    ADOConnection.Execute (insertSQL)
    
    ADOConnection.Close
errorHandler:
    'ないもしない
    If Err.Number <> 0 Then
        errorMsg = errorMsg & "エラー番号：" & Err.Number & vbNewLine
        errorMsg = errorMsg & "エラー内容：" & Err.Description & vbNewLine
        errorMsg = errorMsg & "ヘルプファイル名：" & Err.HelpContext & vbNewLine
        errorMsg = errorMsg & "プロジェクト名：" & Err.Source & vbNewLine
        MsgBox errorMsg
    End If
    
End Function
