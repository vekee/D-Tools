Attribute VB_Name = "ReleaseTools"
'資材リリース用ツール
Sub ExtractFiles()
    
    '全体定数を利用する
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    

    '出力フォルダー
    outputRootFolder = GetSaveDir & "\" & "資材抽出結果_" & Format(Now, "yyyymmddHHMMSS")

    For i = 2 To maxRowNo
        
        On Error GoTo errorHandler
        'ルートDIR
        rootDir = normalizePath(ActiveSheet.Cells(i, 1).value)
        'ファイルパス
        filePath = normalizePath(ActiveSheet.Cells(i, 2).value)
        
        'ファイルの全パス
        fileAllPath = rootDir & "\" & filePath
        
        
        'サブフォルダー作成
        subFolder = Left(filePath, InStrRev(filePath, "\"))
        MkDir outputRootFolder & "\" & subFolder

        
        '対象資材を出力フォルダーにコピーする
        FileCopy fileAllPath, outputRootFolder & "\" & filePath
        
        ActiveSheet.Cells(i, 3).value = "OK"
        
errorHandler:
    If Err.Number <> 0 Then
        ActiveSheet.Cells(i, 3).value = "NG" & "：" & Err.Description
    End If

    Next
    
    MsgBox "抽出完了" & vbCrLf & outputRootFolder
    
End Sub
 
Private Function normalizePath(paramPath As String) As String
               
        paramPath = Replace(paramPath, "/", "\")
        
        If Left(paramPath, 1) = "\" Then
            paramPath = Right(paramPath, Len(paramPath) - 1)
        End If
        
        If Right(paramPath, 1) = "\" Then
            paramPath = Left(paramPath, Len(paramPath) - 1)
        End If
        
        normalizePath = paramPath
        
End Function
 
 
