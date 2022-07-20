Attribute VB_Name = "FormInit"
Public Const g_cnsTITLE = "D-Tools"
Public Sub Workbook_open()
    MsgBox "This code ran at Excel start!"
End Sub


'***********************************************************************************************************************
' �@�\   : �c�[���o�[�������@�\
' �T�v   : �G�N�Z�����J���������A�c�[���o�[�ŁA�{�^�����쐬����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Sub Auto_Open()
    Dim DdataName As Variant
    DdataName = "DTools"
    
    'D-data�𑶍݂��Ă��邱�Ƃ��m�F����
    For Each cb In Application.CommandBars
        If cb.name = DdataName Then
            Exit Sub
        End If
    Next
    
    Dim objTlBar As CommandBar
    Dim objTlBarBtn As CommandBarButton

    Set objTlBar = Application.CommandBars.Add(name:=DdataName, Position:=msoBarTop)
    Set objTlBarBtn = objTlBar.Controls.Add(Type:=msoControlButton)
    
    objTlBarBtn.Style = msoButtonCaption
    objTlBarBtn.Caption = DdataName
    objTlBarBtn.OnAction = "InitForm"
    
    objTlBar.Visible = True
    objTlBar.Protection = msoBarNoChangeVisible
    
    Set objTlBarBtn = Nothing
    Set objTlBar = Nothing
    
    '�o�[�W�����`�F�b�N
    'Call CheckDdataVersion

End Sub
'***********************************************************************************************************************
' �@�\   : �c�[���o�[�������@�\
' �T�v   : �G�N�Z����������A�c�[���o�[�ŁA�{�^�����폜����
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Sub Auto_Close()
    
    Dim DdataName As Variant
    'DdataName = "D-Tools"
    DdataName = "DTools"
    
    '����O�ɁA���[�U�[�̐ݒ�̏���ۑ�����B
    'ThisWorkbook.Save
    
    For Each cb In Application.CommandBars
        If cb.name = DdataName Then
            Dim objBar As CommandBar
            Set objBar = Application.CommandBars(DdataName)
            
            If Not (objBar Is Nothing) Then
                objBar.Delete
                Set objBar = Nothing
            End If
            
            Exit Sub
            
        End If
    Next
    
End Sub
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : Ddata��ʂ�\������
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Sub InitForm()
    DTools.Show
End Sub
'***********************************************************************************************************************
' �@�\   : ���ʋ@�\
' �T�v   : Ddata��ʂ���邷��
' ����   : ��
' �߂�l : ��
'***********************************************************************************************************************
Public Sub CloseForm()
    '����O�ɁA���[�U�[�̐ݒ�̏���ۑ�����B
    'ThisWorkbook.Save
    Unload DTools
End Sub

