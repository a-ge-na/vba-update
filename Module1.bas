Attribute VB_Name = "Module1"
'
'   標準モジュール１
'   既存と更新の両方に存在して、アップデート対象のモジュール
'
Option Explicit

Public Sub dummy1()
    MsgBox "更新前１"
End Sub


Public Sub test()
    Call dummy1     'アップデート対象外のため動作が変わらない
    Call dummy2     'アップデート対象のため前後で動作が変わる
    'Call dummy3    '存在しないため呼び出すとエラー
End Sub

Public Sub update()
    Const UPDATEFILE = "update.zip"
    Const TEMPPREFIX = "goodby__"
    Dim FSO         As New FileSystemObject
    Dim Shell32     As New Shell32.Shell
    Dim Folder      As Shell32.Folder
    Dim File        As Shell32.FolderItem
    Dim TempFolder  As Shell32.Folder
    Dim UpdatePath  As String
    Dim TempPath    As String
    Dim TempFile    As String
    Dim VBProject   As VBProject
    Dim VBComponent As VBComponent
    Dim FileName    As String
    
    'ブックと同一パスにアップデートファイル(zip)があるかを確認
    UpdatePath = FSO.BuildPath(ThisWorkbook.Path, UPDATEFILE)
    If Not FSO.FileExists(UpdatePath) Then
        MsgBox "アップデートファイルが見つかりません。", vbInformation
        Exit Sub
    End If

    'アップデートファイルを展開
    TempPath = FSO.BuildPath(FSO.GetSpecialFolder(TemporaryFolder), FSO.GetTempName())
    FSO.CreateFolder TempPath
    Set TempFolder = Shell32.Namespace(TempPath)
    Set Folder = Shell32.Namespace(UpdatePath)
    For Each File In Folder.Items
        TempFolder.CopyHere File
    Next
    
    'ブック内のプロジェクトと同名のファイルがあれば更新
    Set VBProject = ThisWorkbook.VBProject
    For Each VBComponent In VBProject.VBComponents
        'モジュールのタイプごとに拡張子を設定
        Select Case VBComponent.Type
        Case vbext_ct_StdModule     '標準モジュール(.bas)
            FileName = VBComponent.Name & ".bas"
        Case vbext_ct_ClassModule   'クラスモジュール(.cls)
            FileName = VBComponent.Name & ".cls"
        Case vbext_ct_MSForm        'フォームモジュール(.frm)
            FileName = VBComponent.Name & ".frm"
        End Select
        '同名のファイルがあれば既存モジュールを改名のうえで取り込む
        TempFile = FSO.BuildPath(TempPath, FileName)
        If FSO.FileExists(TempFile) Then
            VBComponent.Name = TEMPPREFIX & VBComponent.Name
            VBProject.VBComponents.Import TempFile
            VBProject.VBComponents.Remove VBComponent
        End If
    Next
    
    '展開先のフォルダを削除
    FSO.DeleteFolder TempPath, True
End Sub


