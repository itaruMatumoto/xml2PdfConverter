Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
    ByVal hwndParent As Long, _
    ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long

Private Declare Function SendMessageAny Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���C���R���g���[���[
' ��P�����i���́j�Fcontrol As IRibbonControl�@���[�U�����삵�����{���̌���
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub MAIN_ON_ACTION(control As IRibbonControl)
    Dim paramObj As New ParamSheetClass

    '�}�N�����s���̉�ʍX�V�����Ȃ�
    Application.ScreenUpdating = False
    '�}�N���̎��s���A���[�U�[�ɓ��͂𑣂����b�Z�[�W��x�����b�Z�[�W��\����}�~����
    Application.DisplayAlerts = False
    ' �����Čv�Z���}�j���A���ɂ���
    Application.Calculation = xlCalculationManual

    ' �e�@�\�̎��s�𔻒肷��
    Select Case control.ID
        Case "Button1"

            ' PDF�ϊ����C���R���g���[�����Ăяo��
            Call Xml2PdfConvertMainController

        Case "Button9"
            ' �o�[�W������\������
            MsgBox "�v���O�����o�[�W������" & paramObj.Version & "�ł��B", vbInformation, ThisWorkbook.name
        
        Case Else
    End Select
    
    Set paramObj = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""
    Application.Calculation = xlCalculationAutomatic

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FPDF�ϊ����C���R���g���[��
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub Xml2PdfConvertMainController()
    Dim rc As Long

    rc = MsgBox("�V�K�ɕϊ����܂����H", vbYesNo + vbQuestion)

    ' �V�K�ϊ��H
    If rc = vbYes Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                ' ���[�g�t�H���_�[�ȉ���T�����āA�t�@�C���p�X���擾����
                Call ExploreFolder(.SelectedItems.Item(1))
            Else
                MsgBox "�L�����Z������܂����B", vbInformation
                GoTo FIN_LABEL
            End If
        End With
    End If

    ' XML->PDF�ɕϊ�����
    Call ConvertTargetPath

FIN_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���[�g�t�H���_�[�ȉ���T�����āA�t�@�C���p�X���擾����
' ��P�����i���́j�Fpath As String  ���[�g�t�H���_�p�X
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ExploreFolder(path As String)
    Dim exObj As New ExploreClass
    Dim helperObj As New HelperClass

    ' ������ݒ肷��
    exObj.SheetName = "xml���X�g"
    exObj.SearchPattern = "^[0-9]{17}[0].xml"      ' �ӕ����i�����t��xml�j�̐��K�\��

    ' �t�H���_��T�����A�����t��xml��T��
    Call exObj.ExploreFolder(helperObj.GetFolderFso(path))

    Set exObj = Nothing
    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�S�p�X��XML->PDF�ɕϊ�����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ConvertTargetPath()
    Dim shObj As New XmlListSheetClass

    shObj.SheetName = "xml���X�g"

    ' �ϊ��Ώۃt�@�C������H
    If shObj.NumOfConvertibleFiles > 0 Then
        ' �ϊ��t�H�[���𗧂��グ��
        Call ShowPdfConvertForm
    Else
        MsgBox "PDF�ϊ��ł���t�@�C��������܂���B", vbExclamation
    End If
    
    Set shObj = Nothing

End Sub


'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�ϊ��t�H�[����\������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ShowPdfConvertForm()

    Load PdfConvertForm
    PdfConvertForm.Show vbModeless

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�A�N�e�B�u�E�C���h�E������ۑ���ʂł��邩���Ď�����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub PollingPrintWindow()
    Dim hWnd As Long
    Dim titles As String * 1000
    Static cnt As Long
    Dim message As String

    Dim debugObj As DebugSheetClass

    Set debugObj = New DebugSheetClass

    Call debugObj.OutPut("�|�[�����O���Ă��܂�...")

    ' �E�C���h�E�n���h�����擾����
    hWnd = FindWindow(vbNullString, "������ʂ𖼑O��t���ĕۑ�")
    
    ' �E�C���h�E�n���h�����擾�ł����H
    If hWnd > 0 Then
        cnt = 0

        Dim hChildWnd As Long

        ' �t�@�C�����̃E�C���h�E�n���h�������߂�
        hChildWnd = FindWindowEx(hWnd, 0, "DUIViewWndClassName", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "DirectUIHWND", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "FloatNotifySink", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "ComboBox", vbNullString)

        ' �t�@�C�����̃E�C���h�E�n���h���ɑ΂��āAPDF�t�@�C���p�X�𑗂�
        Call SendMessageAny(hChildWnd, WM_SETTEXT, 0, PdfConvertForm.targetPdfPath)

        ' �ۑ�(&S)�̃E�C���h�E�n���h�������߂�
        hChildWnd = FindWindowEx(hWnd, 0, "Button", "�ۑ�(&S)")

        ' �ۑ�(&S)�̃E�C���h�E�n���h���ɑ΂��āA���^�[���L�[�𑗂�
        Call PostMessage(hChildWnd, WM_KEYDOWN, VK_RETURN, 0)

    Else
        cnt = cnt + 1

        message = "PDF�ɕϊ����Ă��܂�..." & PdfConvertForm.Progress & " - �Ď����i" & cnt & "��j"

        Call PdfConvertForm.SetMessageTextBox(message)

        Application.OnTime Now + TimeSerial(0, 0, 5), "PollingPrintWindow"
    End If

    Set debugObj = Nothing

End Sub
