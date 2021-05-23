VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "XmlListSheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �v���p�e�B�ϐ�
Private Counter_ As Long
Private SheetName_ As String
Private MaxRowNumber_ As Long
Private NumOfConvertibleFiles_ As Long      ' �ϊ��\�t�@�C����

' �v���C�x�[�g�ϐ�
Private helperObj As HelperClass
Private timeStampObj As TimeStampClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set helperObj = New HelperClass
    Set timeStampObj = New TimeStampClass

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set helperObj = Nothing
    Set timeStampObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�N���A����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub Clear()

    MaxRowNumber_ = 0
    NumOfConvertibleFiles_ = 0
    Counter_ = 0

    With ThisWorkbook.Worksheets(SheetName_)
        .Cells.Clear
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�@�C���p�X���������ށB
' ��P�����i���́j�FfilePath As String  �t�@�C���p�X
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub Writer(filePath As String)
        
    Counter_ = Counter_ + 1

    With ThisWorkbook.Worksheets(SheetName_)
        .Cells(Counter_, "A") = filePath
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�@�C���p�X��ǂݍ��ށB
'                   �g���q�t�B���^�[���ݒ肳��Ă���ꍇ�A���Y�g���q�̃t�@�C���p�X��Ԃ��A
'                   ����ȊO��ǂݔ�΂��B
' �Ԃ�l          �F�t�@�C���p�X�iString�j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function Reader() As String
    Dim num As Long
    Dim filePath As String
    
    Reader = ""

    For num = Counter_ + 1 To MaxRowNumber
        With ThisWorkbook.Worksheets(SheetName_)
            Counter_ = num
            filePath = .Cells(Counter_, "A")
            
            ' �p�X�����݁H
            If helperObj.IsFileExist(filePath) Then
                ' B�񂪋�i�������̃t�@�C���p�X�j�H
                If .Cells(Counter_, "B") = "" Then
                    Reader = filePath
                    ' �t�@�C���p�X���擾�ł����̂Ń��[�v�𔲂���
                    Exit For
                End If
            Else
                Call LogOnSheet("�t�@�C���p�X��������܂���B")
            End If
        End With
    Next num
    
End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�J�����g�s�̉E�[�Z���Ƀ��O���L�^����
' ��P�����i���́j�Fcontents As String  ���e
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub LogOnSheet(contents As String)
    Dim col As Long

    With ThisWorkbook.Worksheets(SheetName_)
        col = .Cells(Counter_, .Columns.Count).End(xlToLeft).Column
        .Cells(Counter_, col + 1) = "[" & timeStampObj.GetTimeStamp & "] " & contents
    End With
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FGet�v���p�e�B�֐�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Get Counter() As Long

    Counter = Counter_

End Property

Public Property Get MaxRowNumber() As Long

    MaxRowNumber = MaxRowNumber_

End Property

Public Property Get NumOfConvertibleFiles() As Long

    NumOfConvertibleFiles = NumOfConvertibleFiles_

End Property

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FLet�v���p�e�B�֐�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Let SheetName(value As String)
    Dim num As Long
    Dim targetFile As File

    SheetName_ = value

    With ThisWorkbook.Worksheets(SheetName_)
        MaxRowNumber_ = .Cells(.Rows.Count, 1).End(xlUp).row
        For num = 1 To MaxRowNumber_
            ' �t�@�C�������݂���H
            If helperObj.IsFileExist(.Cells(num, "A")) Then
                ' B�񂪋�i�������j�H
                If .Cells(num, "B") = "" Then
                    
                    ' �ϊ��\�t�@�C���Ƃ��ăJ�E���g����
                    NumOfConvertibleFiles_ = NumOfConvertibleFiles_ + 1
                    
                    ' ���Y�s��B��ȍ~�̓��e��S�ď���
                    .Range(.Cells(num, "B"), .Cells(num, .Columns.Count)).Clear
                End If
            End If
        Next num
    End With

End Property