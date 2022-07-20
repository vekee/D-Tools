Attribute VB_Name = "Dictionary"

'***********************************************************************************************************************
' 機能   : 辞書登録機能
' 概要   : 入力された情報より、登録した情報を取得して、エクセルに表示する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Function SearchFromDic()

    On Error GoTo errorHandler

    'D-Tools画面の初期設定を実施する
    Load DTools

    Dim selectSQL As String
    selectSQL = "SELECT [論理名],[物理名],[備考],[追加者],[追加日],[削除フラグ] FROM [論物変換テーブル] "
    Dim selectSQLConditions  As String
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    For i = 2 To rowCount
        Dim selectSQLCondition  As String
        selectSQLCondition = ""
        If ActiveSheet.Cells(i, 1).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[論理名] like '%" & ActiveSheet.Cells(i, 1).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 2).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[物理名] like '%" & ActiveSheet.Cells(i, 2).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 3).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[備考] like '%" & ActiveSheet.Cells(i, 3).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 4).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[追加者] like '%" & ActiveSheet.Cells(i, 4).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 5).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[追加日] like '%" & ActiveSheet.Cells(i, 5).value & "%' AND "
        End If
        If ActiveSheet.Cells(i, 6).value <> "" Then
            selectSQLCondition = selectSQLCondition & "[削除フラグ] like '%" & ActiveSheet.Cells(i, 6).value & "%' AND "
        End If

        If selectSQLCondition <> "" Then
            selectSQLCondition = Mid(selectSQLCondition, 1, Len(selectSQLCondition) - 4)
            selectSQLCondition = "(" & selectSQLCondition & ") OR "
            
            selectSQLConditions = selectSQLConditions & selectSQLCondition
        End If
        
    Next
    
    If selectSQLConditions <> "" Then
        selectSQLConditions = Mid(selectSQLConditions, 1, Len(selectSQLConditions) - 3)
        selectSQL = selectSQL & " WHERE " & selectSQLConditions
    End If
    
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Set ADOConnection = connAccessDB()
    '検索SQLを実行する
    ADORecordset.Open selectSQL, ADOConnection
    
    i = 2
    Do Until ADORecordset.EOF
        Set resultFields = ADORecordset.Fields
        ActiveSheet.Cells(i, 1).value = resultFields("論理名").value
        ActiveSheet.Cells(i, 2).value = resultFields("物理名").value
        ActiveSheet.Cells(i, 3).value = resultFields("備考").value
        ActiveSheet.Cells(i, 4).value = resultFields("追加者").value
        ActiveSheet.Cells(i, 5).value = resultFields("追加日").value
        ActiveSheet.Cells(i, 6).value = resultFields("削除フラグ").value
        i = i + 1
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    'G列をクリアする
    ActiveSheet.columns("G").Clear
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("SearchFromDic")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If
    
End Function

'***********************************************************************************************************************
' 機能   : 辞書登録機能
' 概要   :  入力された情報を論物辞書に登録する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Function RegisterToDic()

On Error GoTo errorHandler

    'D-Tools画面の初期設定を実施する
    Load DTools
    
    Dim selectSQL As String
    Dim updateSQL As String
    Dim insertSQL As String
    Dim deleteSQL As String
    Dim sqlConditions  As String
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Set ADOConnection = connAccessDB()
    
    'G列をクリアする
    ActiveSheet.columns("G").Clear
    
    For i = 2 To rowCount
           
            論理名 = ActiveSheet.Cells(i, 1).value
            物理名 = ActiveSheet.Cells(i, 2).value
            備考 = ActiveSheet.Cells(i, 3).value
            追加者 = ActiveSheet.Cells(i, 4).value
            追加日 = ActiveSheet.Cells(i, 5).value
            削除フラグ = ActiveSheet.Cells(i, 6).value
            
            If 論理名 <> "" And 物理名 <> "" And 追加者 <> "" And 追加日 <> "" And 削除フラグ <> "" Then

                selectSQL = "SELECT * FROM [論物変換テーブル] WHERE [削除フラグ] = '0' AND ([論理名] = '" & 論理名 & "' AND [物理名] = '" & 物理名 & "')"
                
                '既存のチェック
                Set ADORecordset = New ADODB.recordset
                ADORecordset.Open selectSQL, ADOConnection
                
                '既存あり状態で、新規登録の場合
                If 削除フラグ = "0" And ADORecordset.EOF = False Then
                    ActiveSheet.Cells(i, 7).value = "論理名と物理名が既に登録されました。"
                    ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                End If
                
                '新規登録の場合
                If 削除フラグ = "0" And ADORecordset.EOF = True Then
                    '既存の辞書より、変換できるかどうかチェックする
                    'ConvertStrInLoop (論理名)
                    insertSQL = "INSERT INTO [論物変換テーブル] ([論理名],[物理名],[備考],[追加者],[追加日],[削除フラグ]) VALUES ('" & 論理名 & "','" & 物理名 & "','" & 備考 & "','" & 追加者 & "','" & 追加日 & "','" & 削除フラグ & "')"
                    ADOConnection.Execute (insertSQL)
                    ActiveSheet.Cells(i, 7).value = "登録済み"
                End If
                
                
                If 削除フラグ = "1" Then
                    selectSQL = "SELECT * FROM [論物変換テーブル] WHERE [論理名] = '" & 論理名 & "' AND [物理名] = '" & 物理名 & "'"
                    Set ADORecordset = New ADODB.recordset
                    ADORecordset.Open selectSQL, ADOConnection
                    '既存に対して論理削除する場合
                    If ADORecordset.EOF = False Then
                        updateSQL = "UPDATE [論物変換テーブル] SET [備考] = '" & 備考 & "',[追加者] = '" & 追加者 & "',[追加日]= '" & 追加日 & "',[削除フラグ] = '" & 削除フラグ & "' WHERE [削除フラグ] = '0' AND ([論理名] = '" & 論理名 & "' AND [物理名] = '" & 物理名 & "')"
                        ADOConnection.Execute (updateSQL)
                        ActiveSheet.Cells(i, 7).value = "論理削除更新済み"
                    Else
                        ActiveSheet.Cells(i, 7).value = "物理削除ができません。対象がリポジトリに存在しない。"
                        ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                    End If
                End If
                
                If 削除フラグ <> "0" And 削除フラグ <> "1" Then
                    selectSQL = "SELECT * FROM [論物変換テーブル] WHERE [論理名] = '" & 論理名 & "' AND [物理名] = '" & 物理名 & "'"
                    Set ADORecordset = New ADODB.recordset
                    ADORecordset.Open selectSQL, ADOConnection
                    '既存に対して物理削除する場合
                    If ADORecordset.EOF = False Then
                        deleteSQL = "DELETE FROM [論物変換テーブル] WHERE [論理名] = '" & 論理名 & "' AND [物理名] = '" & 物理名 & "'"
                        ADOConnection.Execute (deleteSQL)
                        ActiveSheet.Cells(i, 7).value = "物理削除済み"
                    Else
                        ActiveSheet.Cells(i, 7).value = "物理削除ができません。対象がリポジトリに存在しない。"
                        ActiveSheet.Cells(i, 7).Font.colorIndex = 3
                    End If
                    
                End If

            Else
                ActiveSheet.Cells(i, 7).value = "登録失敗！論理名、物理名、追加者、追加日、削除フラグが必要のため、設定してください。"
                ActiveSheet.Cells(i, 7).Font.colorIndex = 3
            End If

    Next
        
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("RegisterToDic")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If
       
    ADORecordset.Close
    ADOConnection.Close
    
End Function

'***********************************************************************************************************************
' 機能   : 論物変換機能
' 概要   : 入力された論理名を物理名へ変換する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Function LogicalToPhysicalByDic()
    On Error GoTo errorHandler
    
    'D-Tools画面の初期設定を実施する
    Load DTools
        
    Dim rowCount As Integer
    rowCount = ActiveSheet.UsedRange.Rows.Count
    
    '全体用変数を初期化する
    PUB_TEMP_VAR_STR = ""
    
    For i = 2 To rowCount
        論理名 = ActiveSheet.Cells(i, 1).value
        ActiveSheet.Cells(i, 2).value = ""
        ActiveSheet.Cells(i, 3).value = ""
        
        
        '論物変換メソッドを呼びだす
        ConvertStrInLoop (論理名)
        
        '変換した内容をエクセルに出力する
        ActiveSheet.Cells(i, 2).value = PUB_TEMP_VAR_STR
        
        If PUB_TEMP_VAR_STR <> "" Then
            '変換できない部分（英字など以外の部分）を再変換する
            PUB_TEMP_VAR_STR = ConvertHiraganaToEnglish(PUB_TEMP_VAR_STR)
            If ActiveSheet.CheckBoxes.value = xlOn Then
                If ActiveSheet.Cells(i, 2).value <> PUB_TEMP_VAR_STR Then
                    ActiveSheet.Cells(i, 3).value = "論物辞書より変換できない文字を存在しています。" & vbCrLf & "【" & PUB_TEMP_VAR_STR & "】を論物辞書に登録しますか"
                End If
            Else
                If ActiveSheet.Cells(i, 2).value <> PUB_TEMP_VAR_STR Then
                    ActiveSheet.Cells(i, 3).value = "論物辞書より変換できない文字を存在しています。"
                End If
            End If
            
        End If
        
        '全体用変数を初期化する
        PUB_TEMP_VAR_STR = ""
    Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("LogicalToPhysicalByDic")
    Else
        'D-Tools画面をクローズする
        Call CloseForm
    End If
    
End Function

'***********************************************************************************************************************
' 機能   : 論物変換機能
' 概要   : 入力された文字列をLoopして、物理名に変換する。
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Function ConvertStrInLoop(convertStr As String)
    Dim RecursionFlag As Boolean
    
    For i = 0 To Len(convertStr) - 1
        Dim subConvertStr As String
        Dim subConvertedStr As String
        subConvertStr = Mid(convertStr, 1, Len(convertStr) - i)
        subConvertedStr = ConvertJPToEnglish(subConvertStr)
        
        '一個目の文字が論物辞書より変換できない場合、かつマスタに定義がある場合
        If Len(subConvertStr) = 1 And ConvertHiraganaToEnglishByMastTab(subConvertStr) <> "" And subConvertedStr = "" Then
            '変更前の文字のままを設定する
            subConvertedStr = subConvertStr
        End If
        
        '変換結果より、再帰変換する
        If subConvertedStr <> "" Then
            '変換できた文字を全体用一時変数に設定する
            PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & Replace(Replace(subConvertedStr, " ", ""), "　", "")
            '変換できた文字を除いて、以外の文字を変更対象とする。
            subConvertStr = Right(convertStr, Len(convertStr) - Len(subConvertStr))
            If subConvertStr <> "" Then
                ConvertStrInLoop (subConvertStr)
            End If
            Exit For
        End If
        
        '一個目の文字が変換できない場合、再帰変換する
        If subConvertedStr = "" And Len(subConvertStr) = 1 Then
            'マスタより、変更する。
            'PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & ConvertHiraganaToEnglish(subConvertStr)
            '変換できない文字も全体用一時変数に設定する
            PUB_TEMP_VAR_STR = PUB_TEMP_VAR_STR & subConvertStr
            '一個目の文字を除いて、以外の文字を変更対象とする。
            subConvertStr = Right(convertStr, Len(convertStr) - Len(subConvertStr))
            If subConvertStr <> "" Then
                ConvertStrInLoop (subConvertStr)
            End If
            Exit For
        End If
    Next

End Function

'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 日本語単語の平仮名から物理名に変換する
' 引数   : String 平仮名
' 戻り値 : String 平仮名の物理名
'***********************************************************************************************************************
Private Function ConvertHiraganaToEnglish(convertStr As String) As String
    Dim ADOConnection As New ADODB.Connection
    Dim ADORecordset As New ADODB.recordset
    Dim 物理名 As String
    Dim convertStrToHiragana As String
    Set ADOConnection = connAccessDB()
    
    For j = 1 To Len(convertStr)
        
        If Mid(convertStr, j, 1) Like "*[a-z,A-Z,0-9,_]" Then
            物理名 = 物理名 & Mid(convertStr, j, 1)
        Else
            convertStrToHiragana = StrConv(Application.GetPhonetic(Mid(convertStr, j, 1)), vbHiragana)
            For i = 1 To Len(convertStrToHiragana)
                Dim hi As String
                Dim convertHIToEN As String
                hi = Mid(convertStrToHiragana, i, 1)
                
                convertHIToEN = ConvertHiraganaToEnglishByMastTab(hi)
                
                物理名 = 物理名 & ConvertHiraganaToEnglishByMastTab(hi)
                
            Next
        End If
    Next
    
    ConvertHiraganaToEnglish = 物理名
    
End Function

'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 日本語単語を平仮名英字マッピングマスタより物理名に変換する
' 引数   : String　日本語単語
' 戻り値 : String　物理名
'***********************************************************************************************************************
Private Function ConvertHiraganaToEnglishByMastTab(convertStr As String) As String
    Dim 物理名 As String
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    
    Dim selectSQL As String
    selectSQL = "SELECT [物理名] FROM [平仮名英字マッピングマスタ] WHERE [削除フラグ] = '0' AND [論理名] = '" & convertStr & "'"
    
    '既存のチェック
    Dim ADORecordset As New ADODB.recordset
    ADORecordset.Open selectSQL, ADOConnection
    
    Do Until ADORecordset.EOF
        Dim resultFields As Fields
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("物理名").value) Then
            物理名 = ""
        Else
            物理名 = resultFields("物理名").value
        End If
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    ConvertHiraganaToEnglishByMastTab = 物理名
End Function
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : 論理名から物理名に変換する
' 引数   : String　論理名
' 戻り値 : String　物理名
'***********************************************************************************************************************
Private Function ConvertJPToEnglish(convertStr As String) As String
    Dim 物理名 As String
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
    
    Dim selectSQL As String
    selectSQL = "SELECT [物理名] FROM [論物変換テーブル] WHERE [削除フラグ] = '0' AND [論理名] = '" & convertStr & "'"
    
    '既存のチェック
    Dim ADORecordset As New ADODB.recordset
    ADORecordset.Open selectSQL, ADOConnection
    
    Do Until ADORecordset.EOF
        Dim resultFields As Fields
        Set resultFields = ADORecordset.Fields
        If IsNull(resultFields("物理名").value) Then
            物理名 = ""
        Else
            物理名 = resultFields("物理名").value
        End If
        
        ADORecordset.MoveNext
    Loop
    
    ADORecordset.Close
    ADOConnection.Close
    
    ConvertJPToEnglish = 物理名
    
End Function
