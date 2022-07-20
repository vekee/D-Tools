Attribute VB_Name = "CodingTools"
Sub ChangeDbColumnNameToDTOVar()

    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    
    For i = 2 To maxRowNo
        ActiveSheet.Cells(i, 2).value = getName(ActiveSheet.Cells(i, 1).value)
    Next
End Sub
Sub CreateDTOByDB()
    
    '全体定数を利用する
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    
    '内部変数作成
    dtoVar = ""
    'set、getメソッド作成
    dtoMethod = ""
    For i = 2 To maxRowNo
        
        dtoType = getType(ActiveSheet.Cells(i, 3).value)
        dtoEngName = getName(ActiveSheet.Cells(i, 2).value)
        dtoJpName = ActiveSheet.Cells(i, 1).value
        dtoGetMethod = "get" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        dtosetMethod = "set" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        
        dtoVar = dtoVar & "    /** " & dtoJpName & " */" & vbCrLf
        dtoVar = dtoVar & "    private " & dtoType & " " & dtoEngName & ";" & vbCrLf
        dtoVar = dtoVar & vbCrLf
        
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を取得する." & vbCrLf
        dtoMethod = dtoMethod & "     * @return " & dtoEngName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public " & dtoType & " " & dtoGetMethod & "() {" & vbCrLf
        If "BigDecimal" = dtoType Then
            dtoMethod = dtoMethod & "        return this." & dtoEngName & " != null ? this." & dtoEngName & " : BigDecimal.ZERO;" & vbCrLf
        Else
            dtoMethod = dtoMethod & "        return this." & dtoEngName & ";" & vbCrLf
        End If
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を設定する." & vbCrLf
        dtoMethod = dtoMethod & "     * @param " & dtoEngName & " " & dtoJpName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public void " & dtosetMethod & "(" & dtoType & " " & dtoEngName & ") {" & vbCrLf
        dtoMethod = dtoMethod & "        this." & dtoEngName & " = " & dtoEngName & ";" & vbCrLf
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        
    Next
    
    
    Dim dtoFile As String
    dtoFile = GetSaveDir & "\" & "DTO自動生成_" & Format(Now, "yyyymmddHHMMSS") & ".txt"
       'ファイルを保存する
    adoStream.WriteText (dtoVar & vbCrLf & dtoMethod)
    adoStream.SaveToFile (dtoFile), adSaveCreateOverWrite
    'ファイルと閉じる
    adoStream.Close
    
    MsgBox "作成完了" & vbCrLf & dtoFile
    
 End Sub
 
Sub CreateCSVDTO()
    
    '全体定数を利用する
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    
    '内部変数作成
    dtoVar = ""
    'set、getメソッド作成
    dtoMethod = ""
    For i = 2 To maxRowNo
        
        dtoType = getType(ActiveSheet.Cells(i, 3).value)
        dtoEngName = getName(ActiveSheet.Cells(i, 2).value)
        dtoJpName = ActiveSheet.Cells(i, 1).value
        dtoSize = ActiveSheet.Cells(i, 4).value
        
        dtoGetMethod = "get" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        dtosetMethod = "set" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        
        dtoVar = dtoVar & "    /** " & dtoJpName & " */" & vbCrLf
        dtoVar = dtoVar & "    @OutputFileColumn(columnIndex = " & i - 2 & ", paddingType = PaddingType.RIGHT, bytes = " & dtoSize & ")" & vbCrLf
        dtoVar = dtoVar & "    private " & dtoType & " " & dtoEngName & ";" & vbCrLf
        dtoVar = dtoVar & vbCrLf
        
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を取得する." & vbCrLf
        dtoMethod = dtoMethod & "     * @return " & dtoEngName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public " & dtoType & " " & dtoGetMethod & "() {" & vbCrLf
        If "BigDecimal" = dtoType Then
            dtoMethod = dtoMethod & "        return this." & dtoEngName & " != null ? this." & dtoEngName & " : BigDecimal.ZERO;" & vbCrLf
        Else
            dtoMethod = dtoMethod & "        return this." & dtoEngName & ";" & vbCrLf
        End If
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を設定する." & vbCrLf
        dtoMethod = dtoMethod & "     * @param " & dtoEngName & " " & dtoJpName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public void " & dtosetMethod & "(" & dtoType & " " & dtoEngName & ") {" & vbCrLf
        dtoMethod = dtoMethod & "        this." & dtoEngName & " = " & dtoEngName & ";" & vbCrLf
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        
    Next
    
    
    Dim dtoFile As String
    dtoFile = GetSaveDir & "\" & "CSV用DTO自動生成_" & Format(Now, "yyyymmddHHMMSS") & ".txt"
       'ファイルを保存する
    adoStream.WriteText (dtoVar & vbCrLf & dtoMethod)
    adoStream.SaveToFile (dtoFile), adSaveCreateOverWrite
    'ファイルと閉じる
    adoStream.Close
    
    MsgBox "作成完了" & vbCrLf & dtoFile
    
 End Sub
 
 Sub CreateSqlMap()
    
    '全体定数を利用する
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    
    'set、getメソッド作成
    sqlMap = ""
    For i = 2 To maxRowNo
        
        dtoJpName = ActiveSheet.Cells(i, 1).value
        dtoEngName = getName(ActiveSheet.Cells(i, 2).value)
        
        sqlMap = sqlMap & "        <!-- " & dtoJpName & " -->" & vbCrLf
        sqlMap = sqlMap & "        <result column=""" & ActiveSheet.Cells(i, 2).value & """ property=""" & dtoEngName & """ />" & vbCrLf

        
    Next
    
    
    Dim sqlMapFile As String
    sqlMapFile = GetSaveDir() & "\" & "SqlMap自動生成_" & Format(Now, "yyyymmddHHMMSS") & ".txt"
       'ファイルを保存する
    adoStream.WriteText (sqlMap)
    adoStream.SaveToFile (sqlMapFile), adSaveCreateOverWrite
    'ファイルと閉じる
    adoStream.Close
    
    MsgBox "作成完了" & vbCrLf & sqlMapFile
    
 End Sub
Sub CreateDTOByUser()
    
    '全体定数を利用する
    Dim adoStream As New ADODB.Stream
    Set adoStream = CreateObject("ADODB.Stream")
    adoStream.Type = adTypeText
    adoStream.Charset = "UTF-8"
    adoStream.LineSeparator = adCRLF
    adoStream.Open
    
    
    maxRowNo = ActiveSheet.UsedRange.Rows.Count
    
    '内部変数作成
    dtoVar = ""
    'set、getメソッド作成
    dtoMethod = ""
    For i = 2 To maxRowNo
        
        dtoType = ActiveSheet.Cells(i, 3).value
        dtoEngName = ActiveSheet.Cells(i, 2).value
        dtoJpName = ActiveSheet.Cells(i, 1).value
        dtoGetMethod = "get" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        dtosetMethod = "set" & UCase(Left(dtoEngName, 1)) & Right(dtoEngName, Len(dtoEngName) - 1)
        
        dtoVar = dtoVar & "    /** " & dtoJpName & " */" & vbCrLf
        dtoVar = dtoVar & "    private " & dtoType & " " & dtoEngName & ";" & vbCrLf
        dtoVar = dtoVar & vbCrLf
        
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を取得する." & vbCrLf
        dtoMethod = dtoMethod & "     * @return " & dtoEngName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public " & dtoType & " " & dtoGetMethod & "() {" & vbCrLf
        If "BigDecimal" = dtoType Then
            dtoMethod = dtoMethod & "        return this." & dtoEngName & " != null ? this." & dtoEngName & " : BigDecimal.ZERO;" & vbCrLf
        Else
            dtoMethod = dtoMethod & "        return this." & dtoEngName & ";" & vbCrLf
        End If
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        dtoMethod = dtoMethod & "    /**" & vbCrLf
        dtoMethod = dtoMethod & "     * " & dtoJpName & "を設定する." & vbCrLf
        dtoMethod = dtoMethod & "     * @param " & dtoEngName & " " & dtoJpName & vbCrLf
        dtoMethod = dtoMethod & "     */" & vbCrLf
        dtoMethod = dtoMethod & "    public void " & dtosetMethod & "(" & dtoType & " " & dtoEngName & ") {" & vbCrLf
        dtoMethod = dtoMethod & "        this." & dtoEngName & " = " & dtoEngName & ";" & vbCrLf
        dtoMethod = dtoMethod & "    }" & vbCrLf
        dtoMethod = dtoMethod & vbCrLf
        
        
    Next
    
    
    Dim dtoFile As String
    dtoFile = GetSaveDir & "\" & "DTO自動生成_" & Format(Now, "yyyymmddHHMMSS") & ".txt"
       'ファイルを保存する
    adoStream.WriteText (dtoVar & vbCrLf & dtoMethod)
    adoStream.SaveToFile (dtoFile), adSaveCreateOverWrite
    'ファイルと閉じる
    adoStream.Close
    
    MsgBox "作成完了" & vbCrLf & dtoFile
    
 End Sub
 
 Private Function getName(dbEngNameParam As String) As String
    dtoEngName = ""
    
    '小文字にする
    dbEngName = LCase(dbEngNameParam)
    
    bigCharFlag = False
    
    For i = 1 To Len(dbEngName)
        If Mid(dbEngName, i, 1) = "_" Then
            bigCharFlag = True
        Else
            If bigCharFlag Then
                dtoEngName = dtoEngName & UCase(Mid(dbEngName, i, 1))
            Else
                dtoEngName = dtoEngName & LCase(Mid(dbEngName, i, 1))
            End If
            bigCharFlag = False
        End If
    Next
    
    If dtoEngName = LCase(dbEngNameParam) Then
        getName = dbEngName
    Else
        getName = dtoEngName
    End If
  
 End Function
 Private Function getType(dbType As String) As String
 
    dtoType = ""
    
    If dbType = "CHAR" Then
        dtoType = "String"
    ElseIf dbType = "NUMBER" Then
        'dtoType = "int"
        dtoType = "BigDecimal"
    ElseIf dbType = "DATE" Then
        dtoType = "Date"
    ElseIf dbType = "VARCHAR2" Or dbType = "VARCHAR" Then
        dtoType = "String"
    ElseIf dbType = "TIMESTAMP" Or aa = "DATETIME" Then
        dtoType = "String"
    ElseIf dbType = "BIGDECIMAL" Or dbType = "DECIMAL" Then
        dtoType = "BigDecimal"
    Else
        dtoType = dbType
    End If
    
    getType = dtoType
    
 End Function
 
 Private Function GetSaveDir() As String
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


