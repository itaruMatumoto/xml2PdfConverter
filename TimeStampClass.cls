VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimeStampClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�^�C���X�^���v�����߂�i�`���P�j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetTimeStamp() As String
    Dim t As SYSTEMTIME
    Dim s As String

    ' ���ݓ������擾����
    Call GetLocalTime(t)

    '// yyyy/mm/dd hh:mm:ss.fff�ɐ��`
    s = Format(t.wYear, "0000")
    s = s & "/"
    s = s & Format(t.wMonth, "00")
    s = s & "/"
    s = s & Format(t.wDay, "00")
    s = s & " "
    s = s & Format(t.wHour, "00")
    s = s & ":"
    s = s & Format(t.wMinute, "00")
    s = s & ":"
    s = s & Format(t.wSecond, "00")
    s = s & "."
    s = s & Format(t.wMilliseconds, "000")

    GetTimeStamp = s

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�^�C���X�^���v�����߂�i�`���Q�j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetTimeStamp2() As String
    Dim t As SYSTEMTIME
    Dim s As String

    ' ���ݓ������擾����
    Call GetLocalTime(t)

    '// yyyy/mm/dd hh:mm:ss.fff�ɐ��`
    s = Format(t.wYear, "0000")
    s = s & Format(t.wMonth, "00")
    s = s & Format(t.wDay, "00")
    s = s & Format(t.wHour, "00")
    s = s & Format(t.wMinute, "00")
    s = s & Format(t.wSecond, "00")
    s = s & Format(t.wMilliseconds, "000")

    GetTimeStamp2 = s

End Function
