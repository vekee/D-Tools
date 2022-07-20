Attribute VB_Name = "SearchDataInExcle"
'***********************************************************************************************************************
' 機能   : エクセルからデータを抽出する機能
' 概要   : 特定セル位置の内容をディレクトリ配下のすべてエクセルから抽出する
' 引数   : Folder 抽出対象ディレクトリ、String 抽出したいの内容
' 戻り値 : 無
'***********************************************************************************************************************
Public Function SearchExcleByContentFromDir(dirFolder As Folder, searchByContent As String)
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
        
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        If dirFolderFile.name Like "*.xls" Or dirFolderFile.name Like "*.xlsx" Then
            On Error GoTo nextFile
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True, PASSWORD:="")
            If dirWorkbook.HasPassword Then
                GoTo nextFile
            End If
            '検索対象ファイルの全シートに検索する
            For Each dirWorksheet In dirWorkbook.Worksheets
                Dim contentStr As Variant
                For Each contentStr In Split(searchByContent, ",")
                    Dim findRange As Range
                    Dim firstRange As Range
                    Set findRange = dirWorksheet.UsedRange.Find(contentStr)
                    If Not findRange Is Nothing Then
                        '見つかった１個目セルを検索結果シートに出力する
                        Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                        Set firstRange = findRange
                        Do
                            Set findRange = dirWorksheet.UsedRange.FindNext(findRange)
                            If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                                Exit Do
                            End If
                            '見つかったセルを検索結果シートに出力する
                            Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                        Loop While firstRange.Address <> findRange.Address
                        
                    End If
                    
                Next

            Next
            
            '検索したエクセルをクロッズする。
            dirWorkbook.Close (False)

        End If
        
nextFile:
    If Err.Number <> 0 Then
        Resume Next
    End If

    Next
    
    For Each subFolder In dirFolder.SubFolders
        Call SearchExcleByContentFromDir(subFolder, searchByContent)
    Next
    
End Function

'***********************************************************************************************************************
' 機能   : エクセルからデータを抽出する機能
' 概要   : 特定セル位置の内容をディレクトリ配下のすべてエクセルから抽出する
' 引数   : Folder 抽出対象ディレクトリ、String 抽出したいのセル位置
' 戻り値 : 無
'***********************************************************************************************************************
Public Function SearchExcleByAddressFromDir(dirFolder As Folder, searchByAddress As String)
    Dim objFSO As FileSystemObject
    Dim subFolder As Folder
    Dim dirFolderFile As File
      
    For Each dirFolderFile In dirFolder.Files
        Dim dirWorkbook As Workbook
        Dim dirWorksheet As Worksheet
        If dirFolderFile.name Like "*.xl*" Then
            On Error GoTo nextFile
            Set dirWorkbook = Workbooks.Open(dirFolderFile.path, UpdateLinks:=False, ReadOnly:=True)
            If dirWorkbook.HasPassword Then
                dirWorkbook.Close (False)
                GoTo nextFile
            End If
            '検索対象ファイルの全シートに検索する
            For Each dirWorksheet In dirWorkbook.Worksheets
                Dim addressStr As Variant
                For Each addressStr In Split(searchByAddress, ",")
                    Dim findRange As Range
                    Dim firstRange As Range
                    
                    Set findRange = dirWorksheet.Range(addressStr)
                    
                    '見つかったセルを検索結果シートに出力する
                    Call WriteSearchResultToSheet(dirFolderFile.path, dirWorksheet.name, findRange.Address, findRange.value)
                                    
                Next
                
            Next

            '検索したエクセルをクロッズする。
            dirWorkbook.Close (False)
nextFile:
        End If
        
    Next
    
    For Each subFolder In dirFolder.SubFolders
        Call SearchExcleByAddressFromDir(subFolder, searchByAddress)
    Next
End Function
'***********************************************************************************************************************
' 機能   : エクセルからデータを抽出する機能
' 概要   : 抽出対象エクセルから抽出したい内容をエクセルに格納する
' 引数   : String 抽出対象エクセルのディレクトリ、String 抽出対象エクセルのシート名、String 抽出対象内容のセル位置、String 抽出対象内容
' 戻り値 : 無
'***********************************************************************************************************************
Private Function WriteSearchResultToSheet(filePath As String, sheetName As String, cellAddress As String, cellValue As String)
    Dim usedRowNo As Integer
    Dim writeToRowNo As Integer
    
    usedRowNo = OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Rows.Count
    writeToRowNo = OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).UsedRange.Rows(usedRowNo).Row + 1
        
    'ヘッダを作成する。
    If usedRowNo = 1 And writeToRowNo = 2 Then
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 1) = "ファイル"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 2) = "シート名"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 3) = "位置"
        OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(1, 4) = "内容"
    End If
        
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 1) = filePath
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 2) = sheetName
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 3) = cellAddress
    OPERATION_WORKBOOK.Sheets(RESULT_SHEET_NAME).Cells(writeToRowNo, 4) = cellValue
    
    '保存
    'OPERATION_WORKBOOK.Save
    
End Function
