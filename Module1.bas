Attribute VB_Name = "Module1"
'
'   �W�����W���[���P
'   �����ƍX�V�̗����ɑ��݂��āA�A�b�v�f�[�g�Ώۂ̃��W���[��
'
Option Explicit

Public Sub dummy1()
    MsgBox "�X�V�O�P"
End Sub


Public Sub test()
    Call dummy1     '�A�b�v�f�[�g�ΏۊO�̂��ߓ��삪�ς��Ȃ�
    Call dummy2     '�A�b�v�f�[�g�Ώۂ̂��ߑO��œ��삪�ς��
    'Call dummy3    '���݂��Ȃ����ߌĂяo���ƃG���[
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
    
    '�u�b�N�Ɠ���p�X�ɃA�b�v�f�[�g�t�@�C��(zip)�����邩���m�F
    UpdatePath = FSO.BuildPath(ThisWorkbook.Path, UPDATEFILE)
    If Not FSO.FileExists(UpdatePath) Then
        MsgBox "�A�b�v�f�[�g�t�@�C����������܂���B", vbInformation
        Exit Sub
    End If

    '�A�b�v�f�[�g�t�@�C����W�J
    TempPath = FSO.BuildPath(FSO.GetSpecialFolder(TemporaryFolder), FSO.GetTempName())
    FSO.CreateFolder TempPath
    Set TempFolder = Shell32.Namespace(TempPath)
    Set Folder = Shell32.Namespace(UpdatePath)
    For Each File In Folder.Items
        TempFolder.CopyHere File
    Next
    
    '�u�b�N���̃v���W�F�N�g�Ɠ����̃t�@�C��������΍X�V
    Set VBProject = ThisWorkbook.VBProject
    For Each VBComponent In VBProject.VBComponents
        '���W���[���̃^�C�v���ƂɊg���q��ݒ�
        Select Case VBComponent.Type
        Case vbext_ct_StdModule     '�W�����W���[��(.bas)
            FileName = VBComponent.Name & ".bas"
        Case vbext_ct_ClassModule   '�N���X���W���[��(.cls)
            FileName = VBComponent.Name & ".cls"
        Case vbext_ct_MSForm        '�t�H�[�����W���[��(.frm)
            FileName = VBComponent.Name & ".frm"
        End Select
        '�����̃t�@�C��������Ί������W���[���������̂����Ŏ�荞��
        TempFile = FSO.BuildPath(TempPath, FileName)
        If FSO.FileExists(TempFile) Then
            VBComponent.Name = TEMPPREFIX & VBComponent.Name
            VBProject.VBComponents.Import TempFile
            VBProject.VBComponents.Remove VBComponent
        End If
    Next
    
    '�W�J��̃t�H���_���폜
    FSO.DeleteFolder TempPath, True
End Sub


