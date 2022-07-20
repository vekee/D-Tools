Attribute VB_Name = "UpVersion"
Public Const ADDIN_FILE_NAME = "D-Tools.xlam"
Public Const UP_VERION_SCRIPT_FILE_NAME = "D-ToolsUpVersion.vbs"

'***********************************************************************************************************************
' 機能   : ツール更新スクリプト出力機能
' 概要   : エクセルを開いった時、ツールバーで、ボタンを作成する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Function CreateUpVersionScript()
    Dim adoObject As Object
    Set adoObject = CreateObject("ADODB.Stream")
    adoObject.Type = adTypeText
    adoObject.Charset = "SJIS"
    adoObject.LineSeparator = adCRLF
    adoObject.Open
    
    adoObject.WriteText ("Option Explicit" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("Dim objFileSys" & vbCrLf)
    adoObject.WriteText ("Dim strFilePathFrom" & vbCrLf)
    adoObject.WriteText ("Dim strFilePathTo" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'3秒スリープ" & vbCrLf)
    adoObject.WriteText ("call WScript.Sleep(3*1000)" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'ファイルシステムを扱うオブジェクトを作成" & vbCrLf)
    adoObject.WriteText ("Set objFileSys = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'コピー元のファイルのパスを指定" & vbCrLf)
    adoObject.WriteText ("strFilePathFrom = ""\\cs.dir.co.jp\document\共通エリア\HN_Project\PJ_BSNS-20121227-00005_M社G新基幹（パートナ共有）\30.ZETTAリアルタイム化・クラウド化\99.WORK\90.JS\98_ツール\D-Tools\" & ADDIN_FILE_NAME & """" & vbCrLf)
    adoObject.WriteText ("strFilePathTo   = " & """" & GetAddinDir & "\" & ADDIN_FILE_NAME & """")
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'エラー発生時にも処理を続行するよう設定" & vbCrLf)
    adoObject.WriteText ("On Error Resume Next" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'ファイルを上書きコピー" & vbCrLf)
    adoObject.WriteText ("Call objFileSys.CopyFile(strFilePathFrom, strFilePathTo)" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'元ファイルが無いなど、エラーになった場合の処理" & vbCrLf)
    adoObject.WriteText ("If Err.Number <> 0 Then" & vbCrLf)
    adoObject.WriteText ("  '何もしない" & vbCrLf)
    adoObject.WriteText ("  'エラー情報をクリアする。" & vbCrLf)
    adoObject.WriteText ("  Err.Clear" & vbCrLf)
    adoObject.WriteText ("End If" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("'「On Error Resume Next」を解除" & vbCrLf)
    adoObject.WriteText ("On Error Goto 0" & vbCrLf)
    adoObject.WriteText (vbCrLf)
    adoObject.WriteText ("Set objFileSys = Nothing " & vbCrLf)
    
    
    Dim vbsFile As String
    vbsFile = Application.UserLibraryPath & "\" & UP_VERION_SCRIPT_FILE_NAME
    
       'ファイルを保存する
    adoObject.SaveToFile (vbsFile), adSaveCreateOverWrite
    'ファイルと閉じる
    adoObject.Close
    
End Function
'***********************************************************************************************************************
' 機能   : エクセルアドインのパス取得
' 概要   :エクセルアドインのパスを取得する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Function GetAddinDir() As String
    GetAddinDir = "C:\Users\" & CreateObject("WScript.Network").UserName & "\AppData\Roaming\Microsoft\AddIns"
End Function
