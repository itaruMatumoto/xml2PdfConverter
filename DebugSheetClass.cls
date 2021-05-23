VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DebugSheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �v���C�x�[�g�ϐ��錾
Private paramObj As ParamSheetClass
Private helperObj As HelperClass
Private timeStampObj As TimeStampClass

Private enableWriteFlag As Boolean

' �萔�錾
Const SheetName As String = "�f�o�b�O"

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()

    Set paramObj = New ParamSheetClass
    Set helperObj = New HelperClass
    Set timeStampObj = New TimeStampClass

    If paramObj.DebugOn Then
        enableWriteFlag = True
        If Not helperObj.IsSheetExist(ThisWorkbook, SheetName) Then
            Call helperObj.CreateSheet(ThisWorkbook, SheetName)
        End If
    Else
        enableWriteFlag = False
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set paramObj = Nothing
    Set helperObj = Nothing
    Set timeStampObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�f�o�b�O���o��
' ��P�����i���́j�Fmessage As String   ���b�Z�[�W
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub OutPut(message As String)
    Dim rowNum As Long
    
    If enableWriteFlag Then
        With ThisWorkbook.Worksheets(SheetName)
            rowNum = .Cells(.Rows.Count, "A").End(xlUp).row
            If Not IsEmpty(.Cells(rowNum, "A")) Then rowNum = rowNum + 1
            .Cells(rowNum, "A") = "[" & timeStampObj.GetTimeStamp & "] " & message
        End With
    End If

End Sub