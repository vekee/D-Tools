Attribute VB_Name = "CONSTS"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�萔
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Const ERROR_INVALID_WINDOW_HANDLE As Long = 1400
Public Const CONN_INFO_SHEET_NAME = "�ڑ����"
Public Const OPERATION_HISTORY_SHEET_NAME = "���엚��"
Public Const VERSION_HISTORY_SHEET_NAME = "�C������"
Public Const INIT_RESULT_SHEET_NAME = "����"
Public Const SELECT_ON = "ON"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�������K�v�̕ϐ�
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'�f�[�^�x�[�X�ڑ�������
Public DB_CONN_INFO_STR As String
Public CONN_INFO_MAX_ROW As Integer

'ActiveWorkBook
Public OPERATION_WORKBOOK As Object
Public RESULT_SHEET_NAME As String

'addInWorkBook
Public addInWorkBook As Object
Public addInConnInfoWS As Object
Public ADDIN_CONN_INFO_WORKSHEET As Object

'DM��`���|�W�g���̃f�B���N�g��
Public DATA_SOURCE_DIR As String

'�Վ��p�S�̕ϐ�
Public PUB_TEMP_VAR_STR As String
Public PUB_TEMP_VAR_OBJ As Object
