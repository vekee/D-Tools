Attribute VB_Name = "FormInit"
Public Const g_cnsTITLE = "D-Tools"
Public Sub Workbook_open()
    MsgBox "This code ran at Excel start!"
End Sub


'***********************************************************************************************************************
' 機能   : ツールバー初期化機能
' 概要   : エクセルを開いった時、ツールバーで、ボタンを作成する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Sub Auto_Open()
    Dim DdataName As Variant
    DdataName = "DTools"
    
    'D-dataを存在していることを確認する
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
    
    'バージョンチェック
    'Call CheckDdataVersion

End Sub
'***********************************************************************************************************************
' 機能   : ツールバー初期化機能
' 概要   : エクセルを閉じた時、ツールバーで、ボタンを削除する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Sub Auto_Close()
    
    Dim DdataName As Variant
    'DdataName = "D-Tools"
    DdataName = "DTools"
    
    '閉じる前に、ユーザーの設定の情報を保存する。
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
' 機能   : 共通機能
' 概要   : Ddata画面を表示する
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Sub InitForm()
    DTools.Show
End Sub
'***********************************************************************************************************************
' 機能   : 共通機能
' 概要   : Ddata画面を閉じるする
' 引数   : 無
' 戻り値 : 無
'***********************************************************************************************************************
Public Sub CloseForm()
    '閉じる前に、ユーザーの設定の情報を保存する。
    'ThisWorkbook.Save
    Unload DTools
End Sub

