VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

' �v���C�x�[�g�ϐ�
Private externalToolPath As String
Private baseFolder As Folder
Private convertedFilesDict As Dictionary
Private signedXmlFile As File

Private helperObj As HelperClass
Private xmlObj As SignedXmlParserClass
Private shellObj As Object
Private debugObj As DebugSheetClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set convertedFilesDict = New Dictionary
    Set helperObj = New HelperClass
    Set xmlObj = New SignedXmlParserClass
    Set shellObj = CreateObject("WScript.Shell")
    Set debugObj = New DebugSheetClass

    externalToolPath = "qpdf.exe"

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set convertedFilesDict = Nothing
    Set helperObj = Nothing
    Set xmlObj = Nothing
    Set shellObj = Nothing
    Set debugObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�[�i�����쐬����
' ��P�����i���́j�FxmlFilePath As String   xml�t�@�C���p�X
' ��Q�����i���́j�FpdfFilePath As String   �ϊ��ς�PDF�t�@�C���p�X
' ��R�����i�o�́j�Frc As Long              ���^�[���R�[�h =0�F������
'                                                          =1�F���l�[������
'                                                          =2�F�}�[�W����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub MakeDeliverables(xmlFilePath As String, pdfFilePath As String, rc As Long)

    rc = 0

    ' �}�[�W��������������
    Call InitializeMergeEnvironment(xmlFilePath)

    ' xml�t�@�C���p�X���L�[�A�ϊ��ς�PDF�t�@�C���p�X��l�Ƃ��āA�����ɓo�^����
    convertedFilesDict.Add xmlFilePath, pdfFilePath

    ' �}�[�W�t�@�C�����쐬�ł��邩�m�F����
    If IsPossibleMerge Then
        Select Case xmlObj.FileCollect.Count
            Case Is = 1
                ' �ϊ��ς�PDF�t�@�C�������l�[������
                Call RenamePdf(pdfFilePath)
                rc = 1
            Case Is > 1
                ' �}�[�W�t�@�C�����쐬����
                Call MergeFiles
                ' �ϊ��ς�PDF��S�č폜����
                Call RemoveConvertedFiles
                rc = 2
            Case Else
                ' �������Ȃ�
        End Select
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FbaseFolder���ݒ莞�AbaseFolder�ύX���Ƀ}�[�W��������������
' ��P�����i���́j�FxmlFilePath As String       xml�t�@�C���p�X
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub InitializeMergeEnvironment(xmlFilePath As String)
    Dim f As File

    Set f = helperObj.GetFileFso(xmlFilePath)

    ' baseFolder���ݒ�H
    If baseFolder Is Nothing Then
        Set baseFolder = helperObj.GetFolderFso(f.ParentFolder)
        ' �����t��xml����͂���
        Call ParseSignedXml
    Else
        ' �x�[�X�t�H���_���ς�����H
        If baseFolder.path <> f.ParentFolder Then
            convertedFilesDict.RemoveAll
            Set baseFolder = helperObj.GetFolderFso(f.ParentFolder)
            ' �����t��xml����͂���
            Call ParseSignedXml
        End If
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�@�C���������l�[������
' ��P�����i���́j�FpdfFilePath As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub RenamePdf(pdfFilePath As String)
    Dim result As Boolean
    Dim targetPdfPath As String
    Dim name As String

    Call debugObj.OutPut("���l�[�����܂�...")

    targetPdfPath = Replace(signedXmlFile.path, ".xml", ".pdf")

    ' �ϊ��ς�PDF�����ɂ���H
    If helperObj.IsFileExist(targetPdfPath) Then
        ' �ϊ��ς�PDF���폜����
        Call helperObj.DeleteFile(targetPdfPath, result)
    End If

    name = Replace(signedXmlFile.name, ".xml", ".pdf")

    ' �ϊ��ς�PDF�t�@�C�������l�[������
    Call helperObj.RenameFile(pdfFilePath, name, result)

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�ϊ��ς�PDF���폜����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub RemoveConvertedFiles()
    Dim var As Variant
    Dim f As File
    Dim num As Long
    
    var = convertedFilesDict.Items
    
    For num = LBound(var) To UBound(var)
        Set f = helperObj.GetFileFso(CStr(var(num)))
        f.Delete
    Next num

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�@�C�����}�[�W����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub MergeFiles()
    Dim command As String
    Dim pdf As String
    Dim f As File

    Call debugObj.OutPut("�}�[�W���܂�...")

    ' �}�[�W���PDF�������߂�
    pdf = Replace(signedXmlFile.path, ".xml", ".pdf")
    
    ' PDF������΍폜����
    If helperObj.IsFileExist(pdf) Then
        Set f = helperObj.GetFileFso(pdf)
        f.Delete
    End If
       
    ' �R�}���h��g�ݗ��Ă�
    command = MakeCommand
    
    Call debugObj.OutPut(command)
    
    ' �R�}���h�����s����
    shellObj.Run command, 0, True
        
    ' �}�[�W���PDF���o����܂ő҂�
    While (Not helperObj.IsFileExist(pdf))
        Sleep 1000
        DoEvents
    Wend

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�R�}���h��g�ݗ��Ă�
' ���Ӂ@qpdf�́A�p�X���ɋ󔒂�����Ɠ����Ȃ��B
'       �R�}���h��g�ݗ��Ă�ۂɃp�X���_�u���N�I�[�e�[�V�����ň͂ޕK�v������B
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function MakeCommand() As String
    Dim pdfAfterMerging As String
    Dim pdfMergeList As String
    Dim eachValue As Variant
    Dim inputPath As String
    Dim eachPath As String

    ' �}�[�W���PDF�������߂�
    pdfAfterMerging = EnclosePathDoubleQuotes(Replace(signedXmlFile.path, ".xml", ".pdf"))
    
    ' �}�[�W���X�g���쐬����
    For Each eachValue In xmlObj.FileCollect
        ' �p�X�ɑg�ݗ��Ă�
        eachPath = baseFolder & "\" & eachValue
        
        ' xml�t�@�C���H
        If helperObj.GetExtensionName(eachPath) = "xml" Then
            If pdfMergeList = "" Then
                pdfMergeList = EnclosePathDoubleQuotes(convertedFilesDict.Item(eachPath))
            Else
                pdfMergeList = pdfMergeList & " " & EnclosePathDoubleQuotes(convertedFilesDict.Item(eachPath))
            End If
        Else
            If pdfMergeList = "" Then
                pdfMergeList = EnclosePathDoubleQuotes(eachPath)
            Else
                pdfMergeList = pdfMergeList & " " & EnclosePathDoubleQuotes(eachPath)
            End If
        End If
        
        ' �P�ڂ̃p�X����̓p�X�Ƃ��Ċo���Ă���
        If inputPath = "" Then
            inputPath = pdfMergeList
        End If

    Next eachValue
    
    ' �R�}���h��g�ݗ��Ă�
    ' �i���Ӂjqpdf�́A�p�X���ɋ󔒂�����Ɠ����Ȃ��B
    '         �R�}���h��g�ݗ��Ă�ۂɃp�X���_�u���N�I�[�e�[�V�����ň͂ޕK�v������B
    MakeCommand = externalToolPath & " " & inputPath & " " & _
                    "--pages " & pdfMergeList & " " & _
                    "-- " & pdfAfterMerging

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�p�X�ɋ󔒂�����P�[�X��z�肵�āA�p�X���_�u���N�I�[�e�[�V�����ň͂�
' ��P�����i���́j�Fpath As String      �t�@�C���p�X
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function EnclosePathDoubleQuotes(path As String) As String

    EnclosePathDoubleQuotes = """" & path & """"

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�ӕ����i�����t��xml�j����͂���
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ParseSignedXml()
    Dim rc As Long
    
    ' �ӕ������擾����
    Set signedXmlFile = GetSignedXml()
    
    ' �ӕ�������͂���
    Call xmlObj.ParseXml(signedXmlFile.path, rc)

    ' ��̓G���[�H
    If rc > 0 Then
        MsgBox "xml��̓G���[���������܂����B" & vbCrLf & signedXmlFile.path, vbCritical
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�ӕ����i�����t��xml�j��Ԃ�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function GetSignedXml() As File
    Dim f As File
    Dim mc As MatchCollection
    
    For Each f In baseFolder.Files
        ' �����t��xml�̐��K�\���p�^�[���Ń}�b�`���O����
        Set mc = helperObj.MatchRegExp(f.name, "^[0-9]{17}[0].xml")
        
        ' �p�^�[������H
        If (mc.Count > 0) Then
            Set GetSignedXml = f
            GoTo FIN_LABEL
        End If
    Next f

FIN_LABEL:

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�}�[�W�t�@�C�����쐬�ł��邩�ۂ���Ԃ�
' �Ԃ�l          �FAs Boolean�@=True:�쐬�ł���^=False:�쐬�ł��Ȃ�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function IsPossibleMerge() As Boolean
    Dim eachValue As Variant
    Dim path As String
    
    IsPossibleMerge = True
    
    For Each eachValue In xmlObj.FileCollect
        path = baseFolder & "\" & eachValue
        ' xml�t�@�C���H
        If helperObj.GetExtensionName(path) = "xml" Then
            ' �ϊ��ς�PDF�͑��݂��Ȃ��H
            If Not convertedFilesDict.Exists(path) Then
                IsPossibleMerge = False
                ' �P�t�@�C���ł�PDF���ϊ��̏ꍇ�A�}�[�W�o���Ȃ��̂Ŕ�����
                Exit For
            End If
        End If
    Next eachValue

End Function
