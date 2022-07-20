Attribute VB_Name = "CONSTS"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'定数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const ERROR_INVALID_WINDOW_HANDLE As Long = 1400
Public Const CONN_INFO_SHEET_NAME = "接続情報"
Public Const OPERATION_HISTORY_SHEET_NAME = "操作履歴"
Public Const VERSION_HISTORY_SHEET_NAME = "修正履歴"
Public Const INIT_RESULT_SHEET_NAME = "結果"
Public Const SELECT_ON = "ON"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'初期化必要の変数
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'データベース接続文字列
Public DB_CONN_INFO_STR As String
Public CONN_INFO_MAX_ROW As Integer

'ActiveWorkBook
Public OPERATION_WORKBOOK As Object
Public RESULT_SHEET_NAME As String

'addInWorkBook
Public addInWorkBook As Object
Public addInConnInfoWS As Object
Public ADDIN_CONN_INFO_WORKSHEET As Object

'DM定義リポジトリのディレクトリ
Public DATA_SOURCE_DIR As String

'臨時用全体変数
Public PUB_TEMP_VAR_STR As String
Public PUB_TEMP_VAR_OBJ As Object
