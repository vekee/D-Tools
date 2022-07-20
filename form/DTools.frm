VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DTools 
   Caption         =   "D-Tools"
   ClientHeight    =   11130
   ClientLeft      =   50
   ClientTop       =   380
   ClientWidth     =   17410
   OleObjectBlob   =   "DTools.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "DTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************************
' 機能   : D-Tools初期化機能
' 概要   : アドインから操作履歴情報を取得して、画面へ設定する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub UserForm_Initialize()

    Set addInWorkBook = Application.ThisWorkbook
    Set OPERATION_WORKBOOK = Application.ActiveWorkbook
    
    '接続情報シート作成
    Set addInConnInfoWS = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME)
    
    '操作履歴シートから内容を設定する
    '実行SQL
    DTools.sqlTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(1, 3).value
    'テーブル物理名
    DTools.GetTableLayoutTableNameTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value
    'InsertSql
    DTools.InsertSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value
    'UpdateSql
    DTools.UpdateSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value
    'DeleteSql
    DTools.DeleteSqlCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value

    'ディレクトリ
    DTools.DirTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(9, 3).value
    '抽出条件
    DTools.GetContentTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(10, 3).value
    '特定の内容
    DTools.GetByContentOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(11, 3).value
    '特定の位置
    DTools.GetByAddressOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(12, 3).value
    'D-Tools定義リポジトリ
    DTools.DMRepositoryTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(13, 3).value
    'テーブル定義書
    DTools.TableCheckBox.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value
    'DM定義成果物格納場所
    DTools.LatestDMDefineFileTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value
    'DB定義
    DTools.TableInfoInDaBaseOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value
    'D-Tools定義リポジトリ
    DTools.TableInfoInDMRepositoryOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value
    

    
    '開始セル
    DTools.SetColorInCellStartTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(24, 3).value
    '終了セル
    DTools.SetColorInCellEndTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(25, 3).value
    '文字開始
    DTools.SetColorInCellCharStartTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(26, 3).value
    '色付け文字数
    DTools.SetColorInCellCharLengTextBox.Text = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(27, 3).value
    '赤
    DTools.SetColorInCellRedColorOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(28, 3).value
    '藍
    DTools.SetColorInCellBlueColorOptionButton.value = addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(29, 3).value
    

    
    'DB接続comboBox作成
    DTools.DBConnInfoComboBox.Clear
    CONN_INFO_MAX_ROW = addInConnInfoWS.UsedRange.Rows.Count
    For i = 1 To CONN_INFO_MAX_ROW
        DTools.DBConnInfoComboBox.AddItem (addInConnInfoWS.Cells(i, 1))
        'TextBoxの初期値を設定する
        If addInConnInfoWS.Cells(i, 3) = SELECT_ON Then
            DTools.DBConnInfoComboBox.ListIndex = i - 1
            DTools.DBConnInfoTextBox.Text = addInConnInfoWS.Cells(i, 2)
        End If
    Next
    
    'フォームをリフレーシュするため
    'Call saveConnInfoButton_Click
    
    '接続情報を初期化
    DB_CONN_INFO_STR = DTools.DBConnInfoTextBox.Text
    '変数初期化
    DATA_SOURCE_DIR = DTools.DMRepositoryTextBox.Text
End Sub
'***********************************************************************************************************************
' 機能   : D-Tools画面を閉じる機能
' 概要   : 「閉じる」ボタンを押下する時、D-Tools画面を閉じる
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub closeFormButton_Click()
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : 接続情報保存機能
' 概要   : 接続情報入力のドロップダウンリストの変更を発生した時、接続情報を更新する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub DBConnInfoComboBox_Change()
    '接続情報を変更する
    CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
    For i = 1 To CONN_INFO_MAX_ROW
        Dim itemValue As String
        itemValue = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1)
        If DTools.DBConnInfoComboBox.Text = itemValue Then
            DTools.DBConnInfoTextBox.Text = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
            addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
            addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON
            
            '全体用変数再設定
            DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
            
        End If
    Next
End Sub
'***********************************************************************************************************************
' 機能   : 接続情報保存機能
' 概要   : 接続情報入力のテキストボックスの変更を発生した時、接続情報保存を自動的に行う
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub DBConnInfoTextBox_Change()
    Call saveConnInfoButton_Click
End Sub

'***********************************************************************************************************************
' 機能   : 接続情報保存機能
' 概要   : 「接続情報保存」ボタンを押下する時、入力した接続情報をアドインに格納する
' 引数   : 無
' 戻り値 : 無
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
            comboBoxSelectedText = "(空白)"
        End If
    
    
        'データ格納
        CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
        
        '既存の接続を変更する場合
        For i = 1 To CONN_INFO_MAX_ROW
            If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1) = comboBoxSelectedText Then
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2) = DBConnInfoTextBox
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 3) = SELECT_ON
                '全体用変数再設定
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                '既存ありフラグ
                existFlag = True
            End If
        Next
        
        '新規の接続を作成する場合
        If existFlag = False Then
            If addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 1) <> "" Then
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 2) = DBConnInfoTextBox
                '全体用変数再設定
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW + 1, 3) = SELECT_ON
            Else
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 1) = comboBoxSelectedText
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 2) = DBConnInfoTextBox
                '全体用変数再設定
                DB_CONN_INFO_STR = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(i, 2)
                
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).columns(3).Clear
                addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Cells(CONN_INFO_MAX_ROW, 3) = SELECT_ON
            End If
        End If

        'ソート
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").Sort Key1:=addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:A"), order1:=xlAscending, Header:=xlNo, MatchCase:=False, SortMethod:=xlPinYin
       
        '重複接続のデータを削除する
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").RemoveDuplicates columns:=Array(1, 2), Header:=xlNo
        
        'comboBox再作成
        DBConnInfoComboBox.Clear
        CONN_INFO_MAX_ROW = addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).UsedRange.Rows.Count
        For i = 1 To CONN_INFO_MAX_ROW
            itemValue = ""
            itemValue = addInConnInfoWS.Cells(i, 1)
            DBConnInfoComboBox.AddItem (itemValue)
            '保存値を表示に設定する
            If comboBoxSelectedText = itemValue Then
                DBConnInfoComboBox.ListIndex = i - 1
            End If
        Next
        
         Set DBConnInfoComboBox = DTools.DBConnInfoComboBox
        
    End If
    
    If comboBoxSelectedText = "" And DBConnInfoTextBox = "" Then
        '接続情報を削除する
        'comboBox再作成
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
        
        '接続リストの一個目を表示する
        DBConnInfoComboBox.ListIndex = 0
        
        'ソート
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").Sort Key1:=addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A1"), order1:=xlAscending, Header:=xlNo, MatchCase:=False, SortMethod:=xlPinYin
        '重複接続のデータを削除する
        addInWorkBook.Worksheets(CONN_INFO_SHEET_NAME).Range("A:B").RemoveDuplicates columns:=Array(1, 2), Header:=xlNo
        
        '変更の接続情報を保存する
        'addInWorkBook.Save
        
    End If
          
End Sub

'***********************************************************************************************************************
' 機能   : SQL実行機能
' 概要   : 「実行」ボタンを押下する時、入力したSQLを実行して、実行結果をエクセルに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub runSQLButton_Click()
    Dim sqls As String
    sqls = DTools.sqlTextBox.Text
    
    '入力情報を保存する
    '実行SQL
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(1, 3).value = sqls
    
    'SQLの空白チェック
    If Replace(Replace(sqls, " ", ""), "　", "") = "" Then
        Exit Sub
    End If
        
    '改行を削除する
    sqls = Replace(sqls, ";" & vbCrLf, ";")
    
    '実行結果の出力先
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '結果出力用シートを作成する
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = Utils.connDB()
    
    On Error GoTo errorHandler
    
    Dim rs As New ADODB.recordset
    rowIndex = 0
    Dim sql As Variant
    For Each sql In Split(sqls, ";")
        'SQLの空白チェック
        If Replace(Replace(sql, " ", ""), "　", "") = "" Then
            GoTo Continue
        End If
    
        colIndex = 1
        rowIndex = rowIndex + 2
        '結果集を初期化する
        
        'sqlの編集する。
        
        Dim recordsAffected As Long
        Set rs = New ADODB.recordset
        Set rs = ADOConnection.Execute(sql, recordsAffected)
        
        If rs.State = 0 Then
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 1).value = sql
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, 2).value = recordsAffected & "件レコードを影響しました。"
            '罫線を付ける
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, 2).Address).Borders.LineStyle = xlContinuous
            GoTo Continue
        End If
        
        Dim resultFields As ADODB.Fields
        Dim resultField As ADODB.field
        Set resultFields = rs.Fields
        
        'カラム名を出力する
        For Each resultField In resultFields
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, colIndex).value = resultField.name
            colIndex = colIndex + 1
        Next
        '罫線を付ける
        resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, colIndex - 1).Address).Borders.LineStyle = xlContinuous
    
        'データを出力する
        Do Until rs.EOF
            rowIndex = rowIndex + 1
            colIndex = 1

            For Each resultField In resultFields
                resultWorkBook.Sheets(RESULT_SHEET_NAME).Cells(rowIndex, colIndex).value = resultField.value
                colIndex = colIndex + 1
            Next
            
            '罫線を付ける
            resultWorkBook.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address & ":" & Cells(rowIndex, colIndex - 1).Address).Borders.LineStyle = xlContinuous
            
            rs.MoveNext
        Loop

Continue:
    Next
    
    If rs.State = 1 Then
        rs.Close
    End If
    ADOConnection.Close
    
    '列の幅自動調整
    resultWorkBook.Sheets(RESULT_SHEET_NAME).UsedRange.columns.AutoFit
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("runSQLButton_Click")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If

End Sub

'***********************************************************************************************************************
' 機能   : 実行計画取得機能
' 概要   : 「実行計画」ボタンを押下する時、入力したSQLの実行結果を取得して、エクセルに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub SqlExecutePlanCommandButton_Click()
    
    '入力情報を保存する
    '実行SQL
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
    
    '①実行計画を解析する
    ADOConnection.Execute explainSql
    
    '②SQLIDを取得する
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
    
    '実行計画を取得する
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
    
    'ユーザーの設定の情報を保存する。
    'addInWorkBook.Save
    
End Sub

'***********************************************************************************************************************
' 機能   : カラム参照機能
' 概要   : 「カラム参照」ボタンを押下する時、入力したテーブル物理名より、テーブルレイアウト情報を取得して、エクセルに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub getTableLayoutButton_Click()
    
    '入力情報を保存する
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value = DTools.GetTableLayoutTableNameTextBox.Text
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value = DTools.TableInfoInDaBaseOptionButton.value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value = DTools.TableInfoInDMRepositoryOptionButton.value
    
    '入力チェック
    If DTools.GetTableLayoutTableNameTextBox.Text = "" Then
        MsgBox "テーブル名を入力してください。"
        Exit Sub
    End If

    'On Error GoTo errorHandler
    
    '結果出力用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    Dim tableName As Variant
    Dim rowIndex As Integer
    rowIndex = 1
    For Each tableName In Split(DTools.GetTableLayoutTableNameTextBox.Text, ",")
        Dim tableNameInfoCollection As New Collection
        
        'テーブル名情報を取得する。
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
                '罫線を付ける
                OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex, 1).Address, Cells(rowIndex, 2).Address).Borders.LineStyle = xlContinuous
                
                'テーブルカラム情報を取得する
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
                
                '罫線を付ける
                 OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(rowIndex + 1, 1).Address & ":" & Cells(rowIndex + 5, colIndex - 1).Address).Borders.LineStyle = xlContinuous
                '列の幅自動調整
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
        'D-Tools画面をクローズする
        Call CloseForm
    End If
    
End Sub
'***********************************************************************************************************************
' 機能   : 試験データ作成機能
' 概要   : 「試験データ作成」ボタンを押下する時、作業用シートを作成して、試験データ作成を行う
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateTestDataCommandButton_Click()
    '入力情報を保存する
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(2, 3).value = DTools.GetTableLayoutTableNameTextBox.Text
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(17, 3).value = DTools.TableInfoInDaBaseOptionButton.value
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(18, 3).value = DTools.TableInfoInDMRepositoryOptionButton.value
       
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools定義リポジトリを設定してください。"
        Exit Sub
    End If
    
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "作成対象テーブル"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(2, 1) = "レコード数"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(3, 1) = "個人識別子" & vbCrLf & "（番号先頭の3～4桁）"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(4, 1) = "番号の枝番" & vbCrLf & "（番号尾部の3桁）"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 1) = "指定論理カラム名" & vbCrLf & "(正式表現記載可)"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 2) = "指定値"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 5) = "利用者名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(5, 6) = "個人識別子" & vbCrLf & "（番号先頭の3～4桁）"

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
        '罫線を付ける
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(usedRowCount, 2).Address).Borders.LineStyle = xlContinuous
    Else
        '罫線を付ける
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(20, 2).Address).Borders.LineStyle = xlContinuous
    End If
    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).Borders.LineStyle = xlContinuous
    
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A5").Interior.Color = RGB(255, 153, 0)
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B5").Interior.Color = RGB(255, 153, 0)
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).Interior.Color = RGB(128, 128, 128)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 45
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A2:A5").RowHeight = 30

    
    '文字の位置を調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").HorizontalAlignment = xlLeft
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").VerticalAlignment = xlTop
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A4").HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:A4").VerticalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(5).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(5).VerticalAlignment = xlCenter
    
    '列の幅を調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 55
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("D1").ColumnWidth = 20
    
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(usedRowCount, 1).Address).columns.AutoFit
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(5, 5).Address & ":" & Cells(dataRowStartIndex - 1, 6).Address).columns.AutoFit

    '「作成する」ボタンを作成する
    With ActiveSheet.Buttons.Add(Range("D1").Left, _
                                 Range("D1").Top, _
                                 Range("D1").Width, _
                                 Range("D1").Height)
        .OnAction = "CreateTestData"
        .Characters.Text = "作成する"
    End With
    
    'D-Tools画面をクローズする
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' 機能   : SQL作成機能
' 概要   : 「SQL作成」ボタンを押下する時、選択されたSQL作成種別より、SQL作成して、ファイルに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateSqlCommandButton_Click()
    
    '入力情報を保存する
    'InsertSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value = DTools.InsertSqlCheckBox.value
    'UpdateSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value = DTools.UpdateSqlCheckBox.value
    'DeleteSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value = DTools.DeleteSqlCheckBox.value
    
    'チェック
    If ActiveSheet.UsedRange.Rows.Count < 7 Then
        MsgBox ("テーブルレイアウト格式が不正！ DToolsで出力したテーブルレイアウトを利用してください！")
    End If
    
    
    '全体定数を利用する
    Set PUB_TEMP_VAR_OBJ = CreateObject("ADODB.Stream")
    PUB_TEMP_VAR_OBJ.Type = adTypeText
    PUB_TEMP_VAR_OBJ.Charset = "UTF-8"
    PUB_TEMP_VAR_OBJ.LineSeparator = adCRLF
    PUB_TEMP_VAR_OBJ.Open
    
    Dim columnsCount As Integer
    columnsCount = ActiveSheet.UsedRange.columns.Count
    Dim usedRowsEndIndex As Integer
    usedRowsEndIndex = ActiveSheet.UsedRange.Rows.Count
    
    'データ情報を配列に格納する
    tableNameJP = ActiveSheet.UsedRange.Cells(1, 1).value
    tableNameEN = ActiveSheet.UsedRange.Cells(1, 2).value
    'スキーマ入力判断
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
    
    'UpdateSql、DeleteSqlを作成する時、Where条件を作成する
    Dim whereSqlSetArray As Variant
    If DTools.UpdateSqlCheckBox.value = True Or DTools.DeleteSqlCheckBox.value = True Then
        whereSqlSetArray = CreateWhereSql()
    End If
    
     'UpdateSqlを作成する
     If DTools.UpdateSqlCheckBox.value = True Then
        Call CreateUpdateSqlSimple(tableNameJP, tableNameEN, columnNameEnArray, dataSetArray, whereSqlSetArray)
     End If
     
     'DeleteSqlを作成する
     If DTools.DeleteSqlCheckBox.value = True Then
        Call CreateDeleteSqlSimple(tableNameJP, tableNameEN, whereSqlSetArray)
     End If
     
     'InsertSqlを作成する
     If DTools.InsertSqlCheckBox.value = True Then
         Call CreateInsertSqlSimple(tableNameJP, tableNameEN, columnNameEnArray, dataSetArray)
     End If
     
    Dim sqlFile As String
    sqlFile = GetSaveDir & "\" & tableNameJP & "_" & Format(Now, "yyyymmddHHMMSS") & ".sql"
       'ファイルを保存する
    PUB_TEMP_VAR_OBJ.SaveToFile (sqlFile), adSaveCreateOverWrite
    'ファイルと閉じる
    PUB_TEMP_VAR_OBJ.Close

    'D-Tools画面をクローズする
    Call CloseForm
    MsgBox "SQLを作成完了しました。" & vbCrLf & "格納場所：" & sqlFile
    
End Sub

'***********************************************************************************************************************
' 機能   : SQL作成機能
' 概要   : 「SQL作成」ボタンを押下する時、選択されたSQL作成種別より、SQL作成して、ファイルに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateSqlCommandButton_Click_BK()
    Dim usedRowsCount As Integer
    Dim usedRowsEndIndex As Integer
    Dim usedRowsStartIndex As Integer
    
    Dim dataRowsStartIndex As Integer


    '入力情報を保存する
    'InsertSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(3, 3).value = DTools.InsertSqlCheckBox.value
    'UpdateSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(4, 3).value = DTools.UpdateSqlCheckBox.value
    'DeleteSql
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(5, 3).value = DTools.DeleteSqlCheckBox.value
    
    
    If DTools.InsertSqlCheckBox.value = False And DTools.UpdateSqlCheckBox.value = False And DTools.DeleteSqlCheckBox.value = False Then
        MsgBox "作成するSql種別を選択してください。"
        Exit Sub
    End If
    
    '全体定数を利用する
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
        '変数初期化
        countinue = False
        oneSetFinish = False
        
        If tableNameEN = "" Then
            If RowIsAllSpace(rowCounter) = True Then
                '次の行を探すに行く
                countinue = True
            Else
                'テーブル物理名を取得する。
                tableNameEN = CreateSql_GetTableName(rowCounter)
            End If
            
            '取得結果のチェック
            If countinue <> True And tableNameEN = "" Then
                MsgBox rowCounter & "行目のデータ格式が不正、ご確認ください。"
                Exit Sub
            Else
                '次の行を探すに行く
                countinue = True
            End If
        End If
        
        'テーブルカラム名を取得する。
        If countinue <> True And tableNameEN <> "" And IsArrayEx(tableColumns) = 0 Then
            If RowIsAllSpace(rowCounter) = True Then
                '次の行を探すに行く
                countinue = True
            Else
                ReDim getTableColumnsResult(2)
                getTableColumnsResult = CreateSql_GetTableColumns(rowCounter)
                
                
                columnStartIndex = getTableColumnsResult(1)
                columnEndIndex = getTableColumnsResult(2)
                
                ReDim tableColumns(columnEndIndex - columnStartIndex + 1)
                tableColumns = getTableColumnsResult(0)
                
            End If
            
            
            '取得結果のチェック
            If countinue <> True And IsArrayEx(tableColumns) = 0 Then
                '次の行を探すに行く
                countinue = True
            End If
            If countinue <> True And IsArrayEx(tableColumns) > 0 Then
                '次の行を探すに行く
                countinue = True
            End If
        
        End If
        
        'データを集める
        If countinue <> True And tableNameEN <> "" And IsArrayEx(tableColumns) = 1 Then
            If RowIsAllSpace(rowCounter) = True Then
                '次の行を探すに行く
                countinue = True
            ElseIf checkTableNameExistInRow(rowCounter) = True Then
                '初期化
                tableNameEN = ""
                Erase getTableColumnsResult
                columnStartIndex = 0
                columnEndIndex = 0
                Erase tableColumns
                Set dataCollection = New Collection
    
                '次の行を探すに行く
                countinue = True
            ElseIf checkStrsExistInRow(rowCounter, Array("CHAR", "VACHAR2", "NUMBER")) <> "0" Then
                'データ型、サイズ、NULL可否を飛ばして、データ行を探して見に行く。
                rowCounter = rowCounter + 2
                '次の行を探すに行く
                countinue = True
            Else
                dataCollection.Add (CreateSql_GetData(rowCounter, columnStartIndex, columnEndIndex))
            End If
            
            '取得結果のチェック
            If countinue <> True And dataCollection.Count = 0 Then
                MsgBox rowCounter & "行目のデータ格式が不正、ご確認ください。"
                Exit Sub
            End If
            
            If countinue <> True And dataCollection.Count > 0 And RowIsAllSpace(rowCounter + 1) = True Then
                '一つセットを設定完了
                oneSetFinish = True
            End If
            
        End If
        
        
        '出力
        If countinue <> True And oneSetFinish = True Then
            
            'テーブルごとの名称を集めいる
            sqlFileName = sqlFileName & tableNameEN & "_"
        
            'InsertSqlを作成する
            If DTools.InsertSqlCheckBox.value = True Then
                Call CreateInsertSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            'UpdateSqlを作成する
            If DTools.UpdateSqlCheckBox.value = True Then
                Call CreateUpdateSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            'DeleteSqlを作成する
        
            If DTools.DeleteSqlCheckBox.value = True Then
                Call CreateDeleteSql(tableNameEN, tableColumns, dataCollection)
            End If
            
            
            '初期化
            tableNameEN = ""
            Erase getTableColumnsResult
            columnStartIndex = 0
            columnEndIndex = 0
            Erase tableColumns
            Set dataCollection = New Collection
    
            '次の行を探すに行く
            countinue = True
           
        End If
            
        rowCounter = rowCounter + 1
        
    Loop


    Dim sqlFile As String
    sqlFile = GetSaveDir & "\" & sqlFileName & Format(Now, "yyyymmdd") & ".sql"
       'ファイルを保存する
    PUB_TEMP_VAR_OBJ.SaveToFile (sqlFile), adSaveCreateOverWrite
    'ファイルと閉じる
    PUB_TEMP_VAR_OBJ.Close
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("CreateSqlCommandButton_Click")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
        MsgBox "SQLを作成完了しました。" & vbCrLf & "格納場所：" & sqlFile
    End If
    
End Sub
'***********************************************************************************************************************
' 機能   : エクセルからデータ抽出機能
' 概要   : 「抽出する」ボタンを押下する時、指定場所のすべてエクセルから指定の内容を抽出する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub GetContentCommandButton_Click()
    Dim objFSO As FileSystemObject
    Dim dirFolder As Folder
    
    Dim dir As String
    dir = DTools.DirTextBox.Text

    '入力情報を保存する
    'ディレクトリ
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(9, 3).value = dir
    '抽出条件
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(10, 3).value = DTools.GetContentTextBox.Text
    '特定の内容
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(11, 3).value = DTools.GetByContentOptionButton.value
    '特定の位置
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(12, 3).value = DTools.GetByAddressOptionButton.value
    
    If dir = "" Then
        MsgBox "検索のディレクトリを入力してください。"
        Exit Sub
    End If
    
    If DTools.GetContentTextBox.Text = "" Then
        MsgBox "検索のデータを入力してください。"
        Exit Sub
    End If
    
    If DTools.GetByContentOptionButton.value = False And DTools.GetByAddressOptionButton.value = False Then
        MsgBox "検索方式(内容、セル位置)を指定してください。"
        Exit Sub
    End If
    
    'On Error GoTo errorHandler
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set dirFolder = objFSO.GetFolder(dir)
    
    '実行結果の出力先
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '結果出力用シートを作成する
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    If DTools.GetByContentOptionButton.value = True Then
        
        Call SearchExcleByContentFromDir(dirFolder, DTools.GetContentTextBox.Text)
        
    ElseIf DTools.GetByAddressOptionButton.value = True Then
    
        Call SearchExcleByAddressFromDir(dirFolder, DTools.GetContentTextBox.Text)
    
    End If
    
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 4).Address).Interior.Color = RGB(255, 153, 0)
    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.columns.AutoFit
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("GetContentCommandButton_Click")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If
    
End Sub
'***********************************************************************************************************************
' 機能   : 章節自動グループ化機能
' 概要   : ワークシートの内容を自動的にグループ化する
' 引数   : 無
' 戻り値 : 無
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
    
    'D-Tools画面をクローズする
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' 機能   : D-Tools定義リポジトリ設定機能
' 概要   : 「ファイル選択」ボタンを押下する時、アクセスファイルの選択画面を起動する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
'
Private Sub FindDMRepositoryCommandButton_Click()
    Dim OpenFileName As String, FileName As String
    OpenFileName = Application.GetOpenFilename("Microsoft Access データベース(*.accdb),*.accdb?")
    If OpenFileName <> "False" Then
        DTools.DMRepositoryTextBox.Text = OpenFileName
    End If
End Sub

'***********************************************************************************************************************
' 機能   : D-Tools定義リポジトリ設定機能
' 概要   : 「保存する」ボタンを押下する時、D-Tools定義リポジトリファイルをアドインに格納する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
'
Private Sub SetDMRepositoryCommandButton_Click()
    '入力情報を保存する
    'D-Tools定義リポジトリ
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(13, 3).value = DTools.DMRepositoryTextBox.Text
    DATA_SOURCE_DIR = DTools.DMRepositoryTextBox.Text
End Sub

'***********************************************************************************************************************
' 機能   : DM同期化機能
' 概要   : 「同期する」ボタンを押下する時、最新のDM定義ファイルから情報を取得して、D-Tools定義リポジトリへ反映する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub SynchronizeDMCommandButton_Click()
   
    '入力情報を保存する
    'テーブル定義書
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value = DTools.TableCheckBox.value
    'DM定義成果物格納場所
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value = DTools.LatestDMDefineFileTextBox.Text
    
    Dim dir As String
    dir = DTools.LatestDMDefineFileTextBox.Text
    
    If dir = "" Then
        MsgBox "D-Tools定義リポジトリを設定してください。"
        Exit Sub
    End If
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools定義リポジトリを指定してください。"
        Exit Sub
    End If
    
    Dim objFSO As New FileSystemObject
    
    If objFSO.FileExists(DATA_SOURCE_DIR) = False Then
        MsgBox "指定のD-Tools定義リポジトリが不存在！"
        Exit Sub
    End If
    
    On Error GoTo result
    
    
    'D-Tools定義リポジトリをバックアップします
    Dim backupDMRepositoryDir As String
    backupDMRepositoryDir = Replace(DATA_SOURCE_DIR, objFSO.getFileName(DATA_SOURCE_DIR), "") & Format(Now, "yyyymmddHHMM") & "_" & objFSO.getFileName(DATA_SOURCE_DIR)
    objFSO.CopyFile DATA_SOURCE_DIR, backupDMRepositoryDir
    
    Dim dirFolder As Folder
    Set dirFolder = objFSO.GetFolder(dir)
    
    Dim result As Boolean
    result = SynchronizeDMDefineInfo(dirFolder)
      
result:
    If Err.Number <> 0 Then
        'D-Tools画面をクローズする
        Call CloseForm
        Call ShowErrorMsg("SynchronizeDMCommandButton_Click")
    ElseIf result = False Then
        'D-Tools画面をクローズする
        'Call CloseForm
    Else
        'D-Tools画面をクローズする
        Call CloseForm
        '完了メッセージを提示する
        MsgBox "同期完了しました。"
    End If
        
End Sub
'***********************************************************************************************************************
' 機能   : テーブル定義書を集約する
' 概要   :
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub ExcleExtracteCommandButton_Click()
       
    '入力情報を保存する
    'テーブル定義書
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(14, 3).value = DTools.TableCheckBox.value
    'DM定義成果物格納場所
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(16, 3).value = DTools.LatestDMDefineFileTextBox.Text
    
    Dim dir As String
    dir = DTools.LatestDMDefineFileTextBox.Text
    
    If dir = "" Then
        MsgBox "D-Tools定義リポジトリを設定してください。"
        Exit Sub
    End If
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools定義リポジトリを指定してください。"
        Exit Sub
    End If
    
    Dim objFSO As New FileSystemObject
    
    If objFSO.FileExists(DATA_SOURCE_DIR) = False Then
        MsgBox "指定のD-Tools定義リポジトリが不存在！"
        Exit Sub
    End If
    
    'On Error GoTo result
    
    Dim tableListSheetName As String
    tableListSheetName = "目次"
    
    '実行結果の出力先
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    
    '結果出力用シートを作成する
    Call createNewSheet(tableListSheetName, resultWorkBook)
    
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 1) = "№"
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 2) = "シート名"
    resultWorkBook.Worksheets(tableListSheetName).Cells(1, 3) = "エクセルファイル名"
    
    Dim dirFolder As Folder
    Set dirFolder = objFSO.GetFolder(dir)
    
    Dim result As Boolean
    'result = ExcleExtracte_sheetCopy(dirFolder)
    result = ExcleExtracteForDic(dirFolder)
    
    fileCounter = OPERATION_WORKBOOK.Sheets(tableListSheetName).UsedRange.Rows.Count
    '罫線を付ける
     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 3).Address).columns.AutoFit
    
result:
    If Err.Number <> 0 Then
        'D-Tools画面をクローズする
        Call CloseForm
        Call ShowErrorMsg("ExcleExtracteCommandButton_Click")
    ElseIf result = False Then
        'D-Tools画面をクローズする
        'Call CloseForm
    Else
        'D-Tools画面をクローズする
        Call CloseForm
        '完了メッセージを提示する
        MsgBox "集約完了しました。"
    End If
End Sub

'***********************************************************************************************************************
' 機能   : 論物辞書登録機能
' 概要   : 「登録シートを作成する」ボタンを押下する時、論物辞書登録用シートを作成して、該シート内で、辞書登録を行う
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateRegisterSheetCommandButton_Click()
        
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools定義リポジトリを設定してください。"
        Exit Sub
    End If
    
    '辞書登録用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "備考"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 4) = "追加者"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 5) = "追加日"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 6) = "削除フラグ"

    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 6).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1:C1").ColumnWidth = 15
    '検索ボタンを作成する
    With ActiveSheet.Buttons.Add(Range("H1").Left, _
                                 Range("H1").Top, _
                                 Range("H1").Width, _
                                 Range("H1").Height)
        .OnAction = "SearchFromDic"
        .Characters.Text = "辞書検索"
    End With
    
    '登録ボタンを作成する
    With ActiveSheet.Buttons.Add(Range("J1").Left, _
                                 Range("J1").Top, _
                                 Range("J1").Width, _
                                 Range("J1").Height)
        .OnAction = "RegisterToDic"
        .Characters.Text = "辞書登録"
    End With
    
    'D-Tools画面をクローズする
    Call CloseForm
    
End Sub

'***********************************************************************************************************************
' 機能   : 論物変換機能
' 概要   : 「論物変換用シートを作成する」ボタンを押下する時、論物変換用シートを作成して、該シート内で、論物変換を行う
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub LogicalToPhysicalCommandButton_Click()
    If DATA_SOURCE_DIR = "" Then
        MsgBox "D-Tools定義リポジトリを設定してください。"
        Exit Sub
    End If
    
    '辞書登録用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "備考"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 40
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "LogicalToPhysicalByDic"
        .Characters.Text = "変換する"
    End With
    
    
    '変換できない時、提示用チェックボックスをを作成する
    With ActiveSheet.CheckBoxes.Add(Range("G1").Left, _
                                   Range("G1").Top, _
                                   Range("G1").Width * 4, _
                                   Range("G1").Height)
        .Characters.Text = "辞書より変換できない場合、自動変換の結果を提示する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : ファイルマージ機能
' 概要   : マージ先ディレクトリ配下にすべてのファイルをマージ元ディレクトリ配下に探して、マージ先ディレクトリへコピーする。
'          マージ元に複数がある場合、最初見つかったファイルをマージ先ディレクトリにコピーする。
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub MergeCommandButton_Click()
        
End Sub

'***********************************************************************************************************************
' 機能   : セル内文字色付け機能
' 概要   : 指定セルに、指定文字数で指定色を付ける
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub SetColorInCellCommandButton_Click()
    '入力情報を保存する
    '開始セル
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(24, 3).value = DTools.SetColorInCellStartTextBox.Text
    '終了セル
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(25, 3).value = DTools.SetColorInCellEndTextBox.Text
    '開始文字
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(26, 3).value = DTools.SetColorInCellCharStartTextBox.Text
    '色付け文字数
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(27, 3).value = DTools.SetColorInCellCharLengTextBox.Text
    '赤
    addInWorkBook.Worksheets(OPERATION_HISTORY_SHEET_NAME).Cells(28, 3).value = DTools.SetColorInCellRedColorOptionButton.value
    '藍
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
    
    'D-Tools画面をクローズする
    Call CloseForm
    
End Sub
'***********************************************************************************************************************
' 機能   : DBカラム名をDTO名に変更する処理
' 概要   : DB物理カラム名をDTO用の変数名に変更する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub DbColumnNameChangeToDtoNameCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DBカラム物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DTO変数名"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "ChangeDbColumnNameToDTOVar"
        .Characters.Text = "変換する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : DBカラム情報よりDTOクラスを作成する
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateDtoClassCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DBカラム論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DBカラム物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "DBカラムの型"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateDTOByDB"
        .Characters.Text = "作成する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : SqlMap用取得項目マッピング作成
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateSqlMapConfigCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DBカラム論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DBカラム物理名"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 2).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateSqlMap"
        .Characters.Text = "作成する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : SqlMap用取得項目マッピング作成
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateCSVDtoClassCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "DBカラム論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "DBカラム物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "DBカラムの型"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateCSVDTO"
        .Characters.Text = "作成する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : ユーザー定義情報よりDTOクラスを作成する
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CreateDtoClassByUserCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "変数論理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "変数物理名"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "変数型"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "CreateDTOByUser"
        .Characters.Text = "作成する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : ユーザー定義情報よりDTOクラスを作成する
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub ExtractFilesToolCommandButton_Click()
    '作業用シートを作成する
    Dim resultWorkBook As Workbook
    Set resultWorkBook = Application.ActiveWorkbook
    Call outPutResultSheet(INIT_RESULT_SHEET_NAME, resultWorkBook)
    
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 1) = "ルートDIR"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 2) = "ファイルパス"
    resultWorkBook.Worksheets(RESULT_SHEET_NAME).Cells(1, 3) = "抽出結果"


    '罫線を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).columns.AutoFit
    '色を付ける
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range(Cells(1, 1).Address & ":" & Cells(1, 3).Address).Interior.Color = RGB(255, 153, 0)
    '行の高さを調整する
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).RowHeight = 20
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).HorizontalAlignment = xlCenter
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Rows(1).VerticalAlignment = xlCenter
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("A1").ColumnWidth = 15
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("B1").ColumnWidth = 80
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("C1").ColumnWidth = 15
    
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Range("E1").ColumnWidth = 10

    '変換するボタンを作成する
    With ActiveSheet.Buttons.Add(Range("E1").Left, _
                                 Range("E1").Top, _
                                 Range("E1").Width, _
                                 Range("E1").Height)
        .OnAction = "ExtractFiles"
        .Characters.Text = "抽出する"
    End With
    
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
'***********************************************************************************************************************
' 機能   : COBOLの可視化番号より詳細設計書の記述漏れがないかをチェックする
' 概要   : なし
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Sub CheckSSCommandButton_Click()
    '設計書チェック機能を初期化する
    'Call AddMenuca
    Call 設計書チェック_手動選択
    
    'D-Tools画面をクローズする
    Call CloseForm
End Sub
