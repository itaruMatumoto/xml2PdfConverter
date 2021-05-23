VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ParamSheetClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' �v���p�e�B�ϐ��錾
Private ExternalToolPath_ As String     ' �O���c�[���p�X
Private DebugOn_ As Boolean
Private Version_ As String

' �v���C�x�[�g�ϐ��錾
Private helperObj As HelperClass

' �萔�錾
Const SheetName As String = "�p�����^"

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    Dim rangeObj As Variant
    Dim num As Long

    Set helperObj = New HelperClass
    DebugOn_ = False

    rangeObj = helperObj.GetRangeVariant(ThisWorkbook.path, SheetName)
    
    For num = LBound(rangeObj) To UBound(rangeObj)
        Select Case rangeObj(num, 1)
            Case "�O���c�[���p�X"
                ExternalToolPath_ = rangeObj(num, 2)
            Case "�v���O�����o�[�W����"
                Version_ = rangeObj(num, 2)
            Case "debug"
                DebugOn_ = True
            Case Else
                ' ���̑�
        End Select
    Next num
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�v���p�e�B�֐�
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Get externalToolPath() As String

    externalToolPath = ExternalToolPath_

End Property

Public Property Get DebugOn() As Boolean

    DebugOn = DebugOn_

End Property

Public Property Get Version() As String

    Version = Version_

End Property