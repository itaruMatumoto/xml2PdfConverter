VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PdfConvertForm 
   Caption         =   "PDF�ϊ�"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10920
   Enabled         =   0   'False
   OleObjectBlob   =   "PdfConvertForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "PdfConvertForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function MessageBoxTimeoutA Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As VbMsgBoxStyle, _
        ByVal wLanguageID As Long, _
        ByVal dwMilliseconds As Long) As Long

' �v���p�e�B�錾
Private targetPdfPath_ As String
Private targetXmlPath_ As String

' �v���C�x�[�g�ϐ�
Private xmlListObj As XmlListSheetClass
Private shellObj As Object
Private helperObj As HelperClass
Private paramObj As ParamSheetClass
Private debugObj As DebugSheetClass
Private timeStampObj As TimeStampClass
Private mergerObj As MergerClass

Private defaultPrinterName As String

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���[�U�t�H�[��������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub UserForm_Initialize()
    Dim list As Variant

    Set xmlListObj = New XmlListSheetClass
    Set shellObj = CreateObject("WScript.Shell")
    Set helperObj = New HelperClass
    Set paramObj = New ParamSheetClass
    Set debugObj = New DebugSheetClass
    Set timeStampObj = New TimeStampClass
    Set mergerObj = New MergerClass

    xmlListObj.SheetName = "xml���X�g"

    targetXmlPath_ = xmlListObj.Reader()

    Call debugObj.OutPut("�u���E�U��\�����܂�...")

    ' �u���E�U��\������
    With Me
        .WebBrowser1.Navigate targetXmlPath_
    End With
    
    ' PDF�o�͂ɐ؂�ւ���
    list = Split(Application.ActivePrinter, " on ")
    defaultPrinterName = list(0)
    
    Call helperObj.ChangeActivePrinter("Microsoft Print to PDF")

FIN_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���[�U�t�H�[�����Ŏ�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub UserForm_Terminate()

    ' �f�t�H���g�v�����^�[�ɐ؂�ւ���
    Call helperObj.ChangeActivePrinter(defaultPrinterName)

    Set xmlListObj = Nothing
    Set shellObj = Nothing
    Set helperObj = Nothing
    Set paramObj = Nothing
    Set debugObj = Nothing
    Set timeStampObj = Nothing
    Set mergerObj = Nothing

    Application.StatusBar = ""

    MsgBox "�v���O�������I�����܂����B", vbInformation

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�h�L�������g�ǂݍ��݊������ɌĂ΂��C�x���g�n���h���[
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim rc As Boolean

    Call debugObj.OutPut("�h�L�������g���ǂݍ��܂�܂���...")

    With Me
        .MessageTextBox = "PDF�ɕϊ����Ă��܂�..." & Progress
    
        ' �ϊ�����PDF�p�X��g�ݗ��Ă�
        targetPdfPath_ = helperObj.GetFileFso(targetXmlPath_).ParentFolder & "\temp_" & timeStampObj.GetTimeStamp2 & ".pdf"

        If helperObj.GetExtensionName(targetXmlPath_) = "xml" Then
        
            Call debugObj.OutPut("��������s���܂�...")

            ' ������s
            ' ��1�����FOLECMDID_PRINT=6              �@[�t�@�C��]���j���[��[���]
            ' ��2�����FOLECMDEXECOPT_DONTPROMPTUSER=2  ���[�U�[�ɓ��͂𑣂����ƂȂ��R�}���h�����s����
            .WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    
        Else
            ' PDF�t�@�C���̃R�s�[���쐬����
            Call helperObj.CopyFile(targetXmlPath_, targetPdfPath_, rc)
        
            ' ���ʕ����쐬����
            Call MakeDeliverables
        End If
    
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FPrintTemplate���������ꂽ���ɌĂ΂��C�x���g�n���h���[
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_PrintTemplateInstantiation(ByVal pDisp As Object)

    Call debugObj.OutPut("PrintTemplate����������܂���...")

    Application.OnTime Now + TimeSerial(0, 0, 5), "PollingPrintWindow"

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FPrintTemplate�����ł������ɌĂ΂��C�x���g�n���h���[
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_PrintTemplateTeardown(ByVal pDisp As Object)


    Call debugObj.OutPut("PrintTemplate�����ł��܂���...")
    Call debugObj.OutPut("PDF�ɕϊ����܂���...")

    Call xmlListObj.LogOnSheet("PDF�ɕϊ����܂����B")

    ' ���ʕ����쐬����
    Call MakeDeliverables

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���ʕ����쐬����B
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub MakeDeliverables()
    Dim rc As Long

    ' �[�i�����쐬����
    Call mergerObj.MakeDeliverables(targetXmlPath_, targetPdfPath_, rc)
    
    ' �쐬�ł����H
    Select Case rc
        Case Is = 1
            Call debugObj.OutPut("PDF�����l�[�����܂���...")
            Call xmlListObj.LogOnSheet("PDF�����l�[�����܂����B")
        Case Is = 2
            Call debugObj.OutPut("���̑�PDF���}�[�W���܂���...")
            Call xmlListObj.LogOnSheet("���̑�PDF���}�[�W���܂����B")
        Case Else
            ' �������Ȃ�
    End Select

    ' �V�[�P���X�𑱂���
    Call ContinueSequence(rc)

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�V�[�P���X�𑱂���
' ��P�����i���́j�Frc As Long          ���ʕ��̏�ԁi>0�F�����A=0�F�r���j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ContinueSequence(rc As Long)
    Dim m As String

    ' ����xml��ǂݍ���
    targetXmlPath_ = xmlListObj.Reader()
    
    ' ����xml������H
    If targetXmlPath_ <> "" Then
        ' ���ʕ����쐬�ł����H
        If rc > 0 Then
            ' �ꎞ��~�𔻒f������
            m = "���~���܂����H" & vbCrLf & "�w����������΁A�R�b��ɍĊJ���܂��B"
            rc = MessageBoxTimeoutA(0&, m, "�m�F", vbYesNo + vbQuestion + vbDefaultButton2, 0&, 3000)
            If rc = vbYes Then
                GoTo UNLOAD_LABEL
            End If
        End If

        Call debugObj.OutPut("�u���E�U��\�����܂�...")

        With Me
            ' �u���E�U��\������
            .WebBrowser1.Navigate targetXmlPath_
        End With
    
        ' �V�[�P���X���p������
        GoTo CONTINUE_LABEL
    
    End If
    
UNLOAD_LABEL:
    ' ����xml�͖������߁A�t�H�[�������i�V�[�P���X��ł��؂�j
    Unload Me

CONTINUE_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���b�Z�[�W��ݒ肷��
' ��P�����i���́j�Fmessage As String       ���b�Z�[�W
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub SetMessageTextBox(message As String)

    With Me
        .MessageTextBox = message
    End With
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FGet�v���p�e�B�֐�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Get targetPdfPath() As String

    targetPdfPath = targetPdfPath_

End Property

Public Property Get Progress() As String

    Progress = "(" & xmlListObj.Counter & "/" & xmlListObj.MaxRowNumber & "�t�@�C��)"

End Property
