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

' プライベート変数宣言
Private paramObj As ParamSheetClass
Private helperObj As HelperClass
Private timeStampObj As TimeStampClass

Private enableWriteFlag As Boolean

' 定数宣言
Const SheetName As String = "デバッグ"

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
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
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set paramObj = Nothing
    Set helperObj = Nothing
    Set timeStampObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：デバッグ情報出力
' 第１引数（入力）：message As String   メッセージ
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
