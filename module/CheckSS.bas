Attribute VB_Name = "CheckSS"
Private Const HEAD_TITLE = "設計書チェック" 'コンテキストメニューのタイトル
Private Const MENU_AUTO = "設計書チェック_自動" 'サブメニューのタイトル（カンマ区切り）
Private Const MENU_MANUAL = "設計書チェック_手動選択" 'サブメニューのタイトル（カンマ区切り）
Private Const MENU_ACT = "設計書チェック" 'サブメニューのSub名（カンマ区切り）

Private Const CHECK_INPUT_SHEET_NAME = "現行調査_セクション構造"
Private Const CHECK_OBJ_SHEET_NAME = "処理内容"

Private Const CHECK_RESULT_SHEET_NAME = "チェック結果"

'***********************************************************************************************************************
' 機能   : 設計書チェック
' 概要   : 可視化資料より、詳細設計書に処理が漏れがないかをチェックして、チェック結果をシートに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Function 設計書チェック_手動選択() As String

    Dim OpenFileName As String
    
    ChDir "\\cs.dir.co.jp\document\共通エリア\HN_Project\PJ_BSNS-20121227-00005_M社G新基幹（パートナ共有）\30.ZETTAリアルタイム化・クラウド化\99.WORK\90.JS\80_検討\18_クラウド化（step2）\03_調査関連\02_設計"
    
    With Application.FileDialog(msoFileDialogFilePicker)
        'ファイルの複数選択を不可能にする
        .AllowMultiSelect = False
        'ファイルフィルタのクリア
        .Filters.Clear
        'ファイルフィルタの追加
        .Filters.Add "エクセルブック", "*.xls*"
        '初期表示フォルダの設定
        .InitialFileName = "\\cs.dir.co.jp\document\共通エリア\HN_Project\PJ_BSNS-20121227-00005_M社G新基幹（パートナ共有）\30.ZETTAリアルタイム化・クラウド化\99.WORK\90.JS\80_検討\18_クラウド化（step2）\03_調査関連\02_設計\"

        If .Show = -1 Then  'ファイルダイアログ表示
            ' [ OK ] ボタンが押された場合
            OpenFileName = .SelectedItems(1)
        Else
           End
        End If
    End With

    Call 設計書チェック(OpenFileName)

End Function
'***********************************************************************************************************************
' 機能   : 設計書チェック
' 概要   : 可視化資料より、詳細設計書に処理が漏れがないかをチェックして、チェック結果をシートに出力する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Private Function 設計書チェック(resultWorkBookPath As String)
    
    On Error GoTo errorHandler
    
    'チェック結果シート作成
    '実行結果の出力先
    Dim resultWorkBook As Workbook
    If resultWorkBookPath <> "" Then
        'Workbooks.Open resultWorkBookPath
        'Set resultWorkBook = Application.ActiveWorkbook
        
        Set resultWorkBook = Workbooks.Open(resultWorkBookPath, UpdateLinks:=False, ReadOnly:=True, PASSWORD:="")
        If resultWorkBook.HasPassword Then
            MsgBox "パスワードを解除してからもう一度試してください！"
        End If
        
    Else
        Set resultWorkBook = Application.ActiveWorkbook
    End If
    
    '結果出力用シートを作成する
    Call createNewSheet(CHECK_RESULT_SHEET_NAME, resultWorkBook)
    
    'チェック対象のセクション構造をチェック結果シートにコピー
    resultWorkBook.Worksheets(CHECK_INPUT_SHEET_NAME).Range("A1").EntireColumn.Copy
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Range("A1").EntireColumn.PasteSpecial (xlPasteAll)
    
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 2).value = "存在個数"
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 3).value = "チェック結果"
    resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(6, 4).value = "備考"
    
    
    'チェック対象を集約する
    Dim checkMaxRowNo As Long
    Dim sectionNo As String
    
    'checkMaxRowNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows.Count
    checkMaxRowNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows(resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).UsedRange.Rows.Count).Row
    
    For i = 7 To checkMaxRowNo
    
        sectionNo = resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 1).value
        sectionNo = SectionNoFilter(sectionNo)
        
        resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 4).value = sectionNo
        
        If sectionNo <> "" Then
        
            Dim findRange As Range
            Dim firstRange As Range
            Dim findCount As Long
            findCount = 0
            Set findRange = resultWorkBook.Worksheets(CHECK_OBJ_SHEET_NAME).UsedRange.Find(sectionNo)
            
            If Not findRange Is Nothing Then
                '見つかった１個目セルを検索結果シートに出力する
                findCount = findCount + 1
                Set firstRange = findRange
                Do
                    Set findRange = resultWorkBook.Worksheets(CHECK_OBJ_SHEET_NAME).UsedRange.FindNext(findRange)
                    If findRange Is Nothing Or firstRange.Address = findRange.Address Then
                        Exit Do
                    End If
                    '見つかったセルを検索結果シートに出力する
                    findCount = findCount + 1
                Loop While firstRange.Address <> findRange.Address
                
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 2).value = findCount
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).value = "存在"
            Else
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 2).value = findCount
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).value = "不存在"
                resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 3).Font.colorIndex = 3
            End If
        Else
            resultWorkBook.Worksheets(CHECK_RESULT_SHEET_NAME).Cells(i, 4).value = "チェック対象外"
        End If
        
    Next
    
   
    'チェック結果シート整形
    '罫線を付ける
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range(Cells(6, 1).Address & ":" & Cells(i - 1, 4).Address).Borders.LineStyle = xlContinuous
    '列の幅自動調整
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range(Cells(6, 1).Address & ":" & Cells(i - 1, 4).Address).columns.AutoFit
    '選択セル
    resultWorkBook.Sheets(CHECK_RESULT_SHEET_NAME).Range("A1").Select
    
    
errorHandler:
    If Err.Number <> 0 Then
        Call ShowErrorMsg("設計書チェック")
    End If
    
End Function

'***********************************************************************************************************************
' 機能   : sectionNoを取り出す
' 概要   : sectionNoを取り出す
' 引数   : sectionNo
' 戻り値 : sectionNo (チェック対象外の場合、空白)
'***********************************************************************************************************************
Private Function SectionNoFilter(sectionNo As String) As String
    sectionNo = Replace(sectionNo, " ", "")
    sectionNo = Replace(sectionNo, "　", "")
    sectionNo = Replace(sectionNo, "*", "")
    sectionNo = Replace(sectionNo, "＊", "")
    sectionNo = Replace(sectionNo, ".", "")
    sectionNo = Replace(sectionNo, "@", "")
    sectionNo = Replace(sectionNo, "＠", "")
    sectionNo = Replace(sectionNo, "┃", "")
    sectionNo = Replace(sectionNo, "┗", "")
    sectionNo = Replace(sectionNo, "━", "")
    sectionNo = Replace(sectionNo, "┣", "")
    sectionNo = Replace(sectionNo, "┣", "")
    
    If sectionNo Like "<*>" Then
        sectionNo = ""
    End If
    
    SectionNoFilter = sectionNo
End Function
'***********************************************************************************************************************
' 機能   : メニューバー追加
' 概要   : 右クリックのメニューバーに追加する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Sub AddMenu()
    On Error Resume Next

    'コントロールに設計書チェックを追加する
    If Not IsControl(HEAD_TITLE) Then
        Set contextmenu = Application.CommandBars("Cell").Controls.Add(Type:=msoControlPopup, Temporary:=True)
        With contextmenu
            .Caption = HEAD_TITLE
            .BeginGroup = False
        End With
    End If

    'サブコントロールに設計書チェックを追加する
    Set contextmenu = Application.CommandBars("Cell").Controls(HEAD_TITLE)
    With contextmenu
        If Not IsSubControl(MENU_AUTO) Then
            With .Controls.Add(Temporary:=True)
                .Caption = MENU_AUTO
                .OnAction = MENU_ACT
                .Enabled = False
            End With
        End If
    End With
    'サブコントロールに設計書チェックを追加する
    With contextmenu
        If Not IsSubControl(MENU_MANUAL) Then
            With .Controls.Add(Temporary:=True)
                .Caption = MENU_MANUAL
                .OnAction = MENU_ACT & "_手動選択"
            End With
        End If
    End With
End Sub
'***********************************************************************************************************************
' 機能   : 内部共通機能
' 概要   : コントロールが存在するかどうをチェックする
' 引数   : String コントロー名
' 戻り値 : Boolean TRUE：存在／FALSE：不存在
'***********************************************************************************************************************
Private Function IsControl(name As String) As Boolean
    Dim found As Boolean

    For Each C In Application.CommandBars("Cell").Controls
        If C.Caption = name Then
            found = True
            Exit For
        End If
    Next C
    IsControl = found
End Function
'***********************************************************************************************************************
' 機能   : 内部共通機能
' 概要   : サブコントロールが存在するかどうをチェックする
' 引数   : String サブコントロー名
' 戻り値 : Boolean TRUE：存在／FALSE：不存在
'***********************************************************************************************************************
Private Function IsSubControl(name As String) As Boolean
    On Error GoTo ex
    Dim found As Boolean
    found = False
    For Each C In Application.CommandBars("Cell").Controls(HEAD_TITLE).Controls
        If C.Caption = name Then
            found = True
            Exit For
        End If
    Next C

ex:
    IsSubControl = found
End Function
