VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExploreClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �v���p�e�B�ϐ�
Private SheetName_ As String            ' �T�����ʂ��L�^����V�[�g��
Private SearchPattern_ As String        ' �T������t�@�C���p�^�[��

' �v���C�x�[�g�ϐ�
Private shObj As XmlListSheetClass
Private xmlObj As SignedXmlParserClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set shObj = New XmlListSheetClass
    Set xmlObj = New SignedXmlParserClass

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Application.StatusBar = ""

    Set shObj = Nothing
    Set xmlObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�H���_��T������
' ��P�����i���́j�Fbase_folder As folder ��t�H���_
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ExploreFolder(baseFolder As Folder)
    Dim eachFolder As Folder
    Dim eachFile As File
    Dim eachDoc As Variant
    Dim mc As MatchCollection
    Dim path As String
    Dim rc As Long

    Dim helperObj As New HelperClass

    Application.StatusBar = "�t�H���_��T�����܂�..." & baseFolder.path

    ' �T�u�t�H���_������H
    If baseFolder.SubFolders.Count > 0 Then
        ' �ċA�I�ɉ��ʊK�w��T������
        For Each eachFolder In baseFolder.SubFolders
            Call ExploreFolder(eachFolder)
        Next eachFolder
    End If
    
    ' ���Y�t�H���_�����̃t�@�C����T��
    For Each eachFile In baseFolder.Files
        Set mc = helperObj.MatchRegExp(eachFile.name, SearchPattern_)
        
        ' �p�^�[���ɊY�������H
        If (mc.Count > 0) Then
            ' xml����͂���
            Call xmlObj.ParseXml(eachFile.path, rc)
            
            ' ��̓G���[�H
            If rc > 0 Then
                MsgBox "xml��̓G���[���������܂����B" & vbCrLf & eachFile.path, vbCritical
                Exit For
            End If
            
            ' �Љ�ی��H
            If xmlObj.SocialInsurance Then
                ' ���Y�t�H���_�ɑ��݂���S�Ă̌����������݂��邩�`�F�b�N����
                For Each eachDoc In xmlObj.FileCollect
                    ' �t�@�C���p�X���擾����
                    path = eachFile.ParentFolder & "\" & eachDoc
                    ' �t�@�C���p�X�����݂��Ȃ��H
                    If Not helperObj.IsFileExist(path) Then
                        MsgBox "���������s�����Ă��܂��B�s�������_�E�����[�h���ĉ������B" & vbCrLf & path, vbCritical
                        GoTo SKIP_LABEL
                    End If
                Next eachDoc
                    
                ' ��������xml���X�g�V�[�g�ɏ����o��
                For Each eachDoc In xmlObj.FileCollect
                    path = eachFile.ParentFolder & "\" & eachDoc
                    ' ���Y�t�H���_�ɂ�xml�����݂���H
                    If xmlObj.XmlFileExist Then
                        If Right(eachDoc, 4) = ".xml" Then
                            ' xml�݂̂��V�[�g�ɏ�������
                            Call shObj.Writer(path)
                        End If
                    Else
                        Call shObj.Writer(path)
                    End If
                Next eachDoc
            End If
            
            ' ���Y�t�H���_�̃`�F�b�N�����������̂ŁA���̃t�H���_�[�Ɉڂ�
            Exit For
        End If
SKIP_LABEL:
    
    Next eachFile

    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FLet�v���p�e�B�֐�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Let SheetName(value As String)

    SheetName_ = value
    
    shObj.SheetName = SheetName_
    shObj.Clear

End Property

Public Property Let SearchPattern(value As String)

    SearchPattern_ = value

End Property