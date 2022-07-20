Attribute VB_Name = "GetDMDefineInfo"
'***********************************************************************************************************************
' 機能   : DM定義同期化機能
' 概要   : 指定ディレクト配下のDM成果物より、DM定義データスースファイルを同期する
' 引数   : Folder DM成果物のディレクトリ、String 同期対象のデータソースファイルのパス
' 戻り値 : True : 同期成功   False : 同期失敗
'***********************************************************************************************************************
Public Function SynchronizeDMDefineInfo(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    
    
    tableDefineSheetName = "テーブル定義書"
    
    'On Error GoTo errorHandler
    
    Dim ADOConnection As New ADODB.Connection
    Set ADOConnection = connAccessDB()
        

    'アウトナンバーを初期化する
    ADOConnection.Execute ("ALTER TABLE [属性定義書] ALTER COLUMN [ID] COUNTER (1,1)")
    'アウトナンバーを初期化する
    ADOConnection.Execute ("ALTER TABLE [テーブル定義書] ALTER COLUMN [ID] COUNTER (1,1)")
    
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
        If dirFolderFile.name Like "*テーブル定義書*.xls*" Then
        
            For Each objSheet In dirWorkbook.Worksheets
                If objSheet.name <> "記入ルール" And objSheet.name <> "変更履歴" And objSheet.name <> "目次" And objSheet.name <> "シーケンス定義" And objSheet.name <> "別紙" Then
                    
                    tableDefineSheetName = objSheet.name
                    
                    Dim 版数 As String
                    Dim 修正日 As String
                    Dim 修正者 As String
                    Dim テーブル名_和名 As String
                    Dim テーブル名_英名 As String
                    Dim テーブル説明 As String
                    Dim insertSql_table As String
                    Dim i As Integer
                    
                    insertSql_table = "INSERT INTO [テーブル定義書] ([テーブル名_和名],[テーブル名_英名],[テーブル説明],[版数],[修正日],[修正者]) VALUES "
                    
        
                    版数 = ""
                    修正日 = ""
                    修正者 = ""
                    
        
                    テーブル名_和名 = dirWorkbook.Worksheets(tableDefineSheetName).Range("C3").value
                    テーブル名_英名 = dirWorkbook.Worksheets(tableDefineSheetName).Range("E3").value
                    テーブル説明 = ""
                    
                    insertSql_table = insertSql_table & "('" & テーブル名_和名 & "'," _
                                                      & "'" & テーブル名_英名 & "'," _
                                                      & "'" & テーブル説明 & "'," _
                                                      & "'" & 版数 & "'," _
                                                      & "'" & 修正日 & "'," _
                                                      & "'" & 修正者 & "')"
                    'テーブル定義書のテーブルをクリアする
                    ADOConnection.Execute ("DELETE FROM [テーブル定義書] WHERE [テーブル名_英名] = '" & テーブル名_英名 & "'")
                    '属性定義書のテーブルをクリアする
                    ADOConnection.Execute ("DELETE FROM [属性定義書] WHERE [テーブル名_英名] = '" & テーブル名_英名 & "'")
                    
                    
                    'テーブル定義書のテーブルへ反映する
                    ADOConnection.Execute (insertSql_table)
                    
        
                    i = 8
                    Do While dirWorkbook.Worksheets(tableDefineSheetName).Cells(i, 4) <> "" And dirWorkbook.Worksheets(tableDefineSheetName).Cells(i, 5) <> ""
                        Dim No As String
                        Dim 属性名_和名 As String
                        Dim カラム名_英名 As String
                        Dim 主キー As String
                        Dim NullAble As String
                        Dim データ型 As String
                        Dim 桁数 As String
                        Dim 小数以下桁数 As String
                        Dim ディフォルト値 As String
                        Dim 旧属性名_和名 As String
                        Dim insertSql_columns As String
                        
                        insertSql_columns = "INSERT INTO [属性定義書] ([テーブル名_英名],[NO],[属性名_和名],[カラム名_英名],[主キー],[NULL],[データ型],[桁数],[小数以下桁数],[ディフォルト値],[旧属性名_和名],[版数],[修正日],[修正者]) VALUES "
                        
                        
                        No = dirWorkbook.Worksheets(tableDefineSheetName).Range("C" & i).value
                        属性名_和名 = dirWorkbook.Worksheets(tableDefineSheetName).Range("D" & i).value
                        カラム名_英名 = dirWorkbook.Worksheets(tableDefineSheetName).Range("E" & i).value
                        主キー = dirWorkbook.Worksheets(tableDefineSheetName).Range("J" & i).value
                        主キー = Replace(主キー, " ", "")
                        主キー = Replace(主キー, "　", "")
                        NullAble = dirWorkbook.Worksheets(tableDefineSheetName).Range("K" & i).value
                        データ型 = dirWorkbook.Worksheets(tableDefineSheetName).Range("F" & i).value
                        桁数 = dirWorkbook.Worksheets(tableDefineSheetName).Range("G" & i).value
                        小数以下桁数 = dirWorkbook.Worksheets(tableDefineSheetName).Range("H" & i).value
                        ディフォルト値 = Replace(dirWorkbook.Worksheets(tableDefineSheetName).Range("L" & i).value, "'", "''")
                        旧属性名_和名 = ""
                        
                        insertSql_columns = insertSql_columns & "('" & テーブル名_英名 & "'," _
                                                              & Val(No) & "," _
                                                              & "'" & 属性名_和名 & "'," _
                                                              & "'" & カラム名_英名 & "'," _
                                                              & "'" & 主キー & "'," _
                                                              & "'" & NullAble & "'," _
                                                              & "'" & データ型 & "'," _
                                                              & Val(桁数) & "," _
                                                              & Val(小数以下桁数) & "," _
                                                              & "'" & ディフォルト値 & "'," _
                                                              & "'" & 旧属性名_和名 & "'," _
                                                              & "'" & 版数 & "'," _
                                                              & "'" & 修正日 & "'," _
                                                              & "'" & 修正者 & "')"
        
                        '属性定義書のテーブルへ反映する
                        ADOConnection.Execute (insertSql_columns)
                        
                        i = i + 1
                    Loop
                
                End If
            Next
        
            
        End If
            
        '検索したエクセルをクロッズする。
        dirWorkbook.Close (False)
        
    Next
    
    'アウトナンバーを初期化する
    ADOConnection.Execute ("ALTER TABLE [属性定義書] ALTER COLUMN [ID] COUNTER (1,1)")
    'アウトナンバーを初期化する
    ADOConnection.Execute ("ALTER TABLE [テーブル定義書] ALTER COLUMN [ID] COUNTER (1,1)")
    
    '再帰処理をしない
    'For Each subFolder In dirFolder.SubFolders
    '    Call SearchExcleByContentFromDir(subFolder, searchByContent)
    'Next
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("SynchronizeDMDefineInfo")
        result = False
    End If
    
    ADOConnection.Close
    
    SynchronizeDMDefineInfo = result
    
End Function
'***********************************************************************************************************************
' 機能   : テーブル定義書集約機能
' 概要   : 指定ディレクト配下のDM成果物より、一つエクセルに集約する機能
' 引数   : Folder DM成果物のディレクトリ、String 同期対象のデータソースファイルのパス
' 戻り値 : True : 同期成功   False : 同期失敗
'***********************************************************************************************************************
Public Function ExcleExtracte(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    fileCounter = 1
    
    tableListSheetName = "目次"
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                If dirWorksheet.name <> "項目説明" And Not (dirWorksheet.name Like "*記入例*") Then
                    
                    'シート名
                    tableDefineSheetName = dirWorksheet.name
                    
                    '目次
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = dirWorksheet.Cells(3, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirWorksheet.Cells(4, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 4).value = dirWorksheet.Cells(2, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 5).value = dirFolderFile.name
        
                    'テーブルシート作成
                    Call createNewSheet(tableDefineSheetName, OPERATION_WORKBOOK)
                    
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 1) = "項番"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 2) = "項目名称"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 3) = "階層"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 4) = "物理名"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 5) = "種別"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 6) = "バイト数"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 7) = "桁数"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 8) = "反復"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 9) = "開始位置"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 10) = "終了位置"
                    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 11) = "説明"
                    
                    '各テーブルシート
                    i = 8
                    Do While dirWorksheet.Cells(i, 2) <> "" Or dirWorksheet.Cells(i, 14) <> "" Or dirWorksheet.Cells(i, 15) <> "" Or dirWorksheet.Cells(i, 29) <> ""
                       
                        
                        項番 = dirWorksheet.Range("A" & i).value
                        項目名称 = dirWorksheet.Range("B" & i).value _
                                 & dirWorksheet.Range("C" & i).value _
                                 & dirWorksheet.Range("D" & i).value _
                                 & dirWorksheet.Range("E" & i).value _
                                 & dirWorksheet.Range("F" & i).value _
                                 & dirWorksheet.Range("G" & i).value _
                                 & dirWorksheet.Range("H" & i).value _
                                 & dirWorksheet.Range("I" & i).value _
                                 & dirWorksheet.Range("J" & i).value _
                                 & dirWorksheet.Range("K" & i).value _
                                 & dirWorksheet.Range("L" & i).value _
                                 & dirWorksheet.Range("M" & i).value
                                 
                        階層 = dirWorksheet.Range("N" & i).value
                        物理名 = dirWorksheet.Range("O" & i).value
                        種別 = dirWorksheet.Range("Z" & i).value
                        バイト数 = dirWorksheet.Range("AC" & i).value & dirWorksheet.Range("AD" & i).value
                        If 種別 = "P" Then
                            桁数 = dirWorksheet.Range("AE" & i).value & "." & dirWorksheet.Range("AF" & i).value & dirWorksheet.Range("AG" & i).value
                        End If
                        反復 = dirWorksheet.Range("AI" & i).value
                        開始位置 = dirWorksheet.Range("AJ" & i).value
                        終了位置 = dirWorksheet.Range("AK" & i).value
                        説明 = dirWorksheet.Range("AM" & i).value
        
        
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 1) = 項番
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 2) = 項目名称
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 3) = 階層
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 4) = 物理名
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 5) = 種別
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 6) = バイト数
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 7) = 桁数
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 8) = 反復
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 9) = 開始位置
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 10) = 終了位置
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(i - 7 + 1, 11) = 説明
                            
                        
                        i = i + 1
                    Loop
                    
                    fileCounter = fileCounter + 1
                    
                    
                    '罫線を付ける
                     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 5).Address).Borders.LineStyle = xlContinuous
                    '列の幅自動調整
                     OPERATION_WORKBOOK.Sheets(tableListSheetName).Range(Cells(1, 1).Address & ":" & Cells(fileCounter, 5).Address).columns.AutoFit
                     
                    '罫線を付ける
                     OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Range(Cells(1, 1).Address & ":" & Cells(i - 7, 11).Address).Borders.LineStyle = xlContinuous
                    '列の幅自動調整
                     OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Range(Cells(1, 1).Address & ":" & Cells(i - 7, 11).Address).columns.AutoFit
                End If
                
            Next
            
            
            '検索したエクセルをクロッズする。
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '再帰処理をする
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracte(subFolder)
    Next
    
    ExcleExtracte = True

    
End Function

'***********************************************************************************************************************
' 機能   : テーブル定義書集約機能
' 概要   : 指定ディレクト配下のDM成果物より、一つエクセルに集約する機能
' 引数   : Folder DM成果物のディレクトリ、String 同期対象のデータソースファイルのパス
' 戻り値 : True : 同期成功   False : 同期失敗
'***********************************************************************************************************************
Public Function ExcleExtracte_sheetCopy(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    
    
    tableListSheetName = "目次"
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                
                    dirWorksheet.Copy After:=OPERATION_WORKBOOK.Worksheets(OPERATION_WORKBOOK.Worksheets.Count)
                    
                    tableDefineSheetName = OPERATION_WORKBOOK.Worksheets(OPERATION_WORKBOOK.Worksheets.Count).name
                    
                    fileCounter = OPERATION_WORKBOOK.Sheets(tableListSheetName).UsedRange.Rows.Count

                    '目次
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = tableDefineSheetName
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirFolderFile.name
                    
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Hyperlinks.Add Anchor:=OPERATION_WORKBOOK.Sheets(tableListSheetName).Range("B" & (fileCounter + 1)), Address:="", SubAddress:="'" & tableDefineSheetName & "'" & "!A1", TextToDisplay:=tableDefineSheetName
                    
                    fileCounter = fileCounter + 1
                
            Next
            
            'コピーした内容を保存する
            'OPERATION_WORKBOOK.Save
            
            '検索したエクセルをクロッズする。
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '再帰処理をする
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracte_sheetCopy(subFolder)
    Next
    
    ExcleExtracte_sheetCopy = True

    
End Function

'***********************************************************************************************************************
' 機能   : テーブル定義書集約機能
' 概要   : 指定ディレクト配下のDM成果物より、一つエクセルに集約する機能
' 引数   : Folder DM成果物のディレクトリ、String 同期対象のデータソースファイルのパス
' 戻り値 : True : 同期成功   False : 同期失敗
'***********************************************************************************************************************
Public Function ExcleExtracteForDic(dirFolder As Folder) As Boolean
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
    Dim tableListSheetName As String
    Dim tableDefineSheetName As String
    Dim result As Boolean
    result = True
    

    Dim fileCounter As Long
    fileCounter = 1
    
    Dim columnCounter As Long
    columnCounter = 1
    
    tableListSheetName = "目次"
    
    'テーブルシート作成
    tableDefineSheetName = "辞書用"
    Call createNewSheet(tableDefineSheetName, OPERATION_WORKBOOK)
    
    
    
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 1) = "項番"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 2) = "項目名称"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 3) = "階層"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 4) = "物理名"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 5) = "種別"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 6) = "バイト数"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 7) = "桁数"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 8) = "反復"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 9) = "開始位置"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 10) = "終了位置"
    OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(1, 11) = "説明"
    
    
    For Each dirFolderFile In dirFolder.Files
    
        If dirFolderFile.name Like "*.xls" Then
        
            Dim dirWorkbook As Workbook
            Dim dirWorksheet As Worksheet
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
                        
            For Each dirWorksheet In dirWorkbook.Worksheets

                If dirWorksheet.name <> "項目説明" And Not (dirWorksheet.name Like "*記入例*") Then
                      
                    '目次
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 1).value = fileCounter
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 2).value = dirWorksheet.Cells(3, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 3).value = dirWorksheet.Cells(4, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 4).value = dirWorksheet.Cells(2, 4).value
                    OPERATION_WORKBOOK.Sheets(tableListSheetName).Cells(fileCounter + 1, 5).value = dirFolderFile.name
                    
                    '各テーブルシート
                    i = 8
                    Do While dirWorksheet.Cells(i, 2) <> "" Or dirWorksheet.Cells(i, 14) <> "" Or dirWorksheet.Cells(i, 15) <> "" Or dirWorksheet.Cells(i, 29) <> ""
                       
                        
                        項番 = dirWorksheet.Range("A" & i).value
                        項目名称 = dirWorksheet.Range("B" & i).value _
                                 & dirWorksheet.Range("C" & i).value _
                                 & dirWorksheet.Range("D" & i).value _
                                 & dirWorksheet.Range("E" & i).value _
                                 & dirWorksheet.Range("F" & i).value _
                                 & dirWorksheet.Range("G" & i).value _
                                 & dirWorksheet.Range("H" & i).value _
                                 & dirWorksheet.Range("I" & i).value _
                                 & dirWorksheet.Range("J" & i).value _
                                 & dirWorksheet.Range("K" & i).value _
                                 & dirWorksheet.Range("L" & i).value _
                                 & dirWorksheet.Range("M" & i).value
                                 
                        階層 = dirWorksheet.Range("N" & i).value
                        物理名 = dirWorksheet.Range("O" & i).value
                        種別 = dirWorksheet.Range("Z" & i).value
                        バイト数 = dirWorksheet.Range("AC" & i).value & dirWorksheet.Range("AD" & i).value
                        If 種別 = "P" Then
                            桁数 = dirWorksheet.Range("AE" & i).value & "." & dirWorksheet.Range("AF" & i).value & dirWorksheet.Range("AG" & i).value
                        End If
                        反復 = dirWorksheet.Range("AI" & i).value
                        開始位置 = dirWorksheet.Range("AJ" & i).value
                        終了位置 = dirWorksheet.Range("AK" & i).value
                        説明 = dirWorksheet.Range("AM" & i).value
        
                        
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 1) = 項番
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 2) = 項目名称
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 3) = 階層
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 4) = 物理名
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 5) = 種別
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 6) = バイト数
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 7) = 桁数
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 8) = 反復
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 9) = 開始位置
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 10) = 終了位置
                        'OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 11) = 説明
                        OPERATION_WORKBOOK.Sheets(tableDefineSheetName).Cells(columnCounter, 11) = dirFolderFile.name
                        
                        
                        i = i + 1
                        columnCounter = columnCounter + 1
                    Loop
                    
                    fileCounter = fileCounter + 1
                    
                    

                End If
                
            Next
            
            
            '検索したエクセルをクロッズする。
            dirWorkbook.Close (False)
        
        End If
        
    Next
    
    
    '再帰処理をする
    For Each subFolder In dirFolder.SubFolders
        Call ExcleExtracteForDic(subFolder)
    Next
    
    ExcleExtracteForDic = True

    
End Function

