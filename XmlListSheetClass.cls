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

' プロパティ変数
Private Counter_ As Long
Private SheetName_ As String
Private MaxRowNumber_ As Long
Private NumOfConvertibleFiles_ As Long      ' 変換可能ファイル数

' プライベート変数
Private helperObj As HelperClass
Private timeStampObj As TimeStampClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set helperObj = New HelperClass
    Set timeStampObj = New TimeStampClass

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set helperObj = Nothing
    Set timeStampObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：クリア処理
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
' 機能　　　　　　：ファイルパスを書き込む。
' 第１引数（入力）：filePath As String  ファイルパス
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub Writer(filePath As String)
        
    Counter_ = Counter_ + 1

    With ThisWorkbook.Worksheets(SheetName_)
        .Cells(Counter_, "A") = filePath
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ファイルパスを読み込む。
'                   拡張子フィルターが設定されている場合、当該拡張子のファイルパスを返し、
'                   それ以外を読み飛ばす。
' 返り値          ：ファイルパス（String）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function Reader() As String
    Dim num As Long
    Dim filePath As String
    
    Reader = ""

    For num = Counter_ + 1 To MaxRowNumber
        With ThisWorkbook.Worksheets(SheetName_)
            Counter_ = num
            filePath = .Cells(Counter_, "A")
            
            ' パスが存在？
            If helperObj.IsFileExist(filePath) Then
                ' B列が空（未処理のファイルパス）？
                If .Cells(Counter_, "B") = "" Then
                    Reader = filePath
                    ' ファイルパスが取得できたのでループを抜ける
                    Exit For
                End If
            Else
                Call LogOnSheet("ファイルパスが見つかりません。")
            End If
        End With
    Next num
    
End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：カレント行の右端セルにログを記録する
' 第１引数（入力）：contents As String  内容
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub LogOnSheet(contents As String)
    Dim col As Long

    With ThisWorkbook.Worksheets(SheetName_)
        col = .Cells(Counter_, .Columns.Count).End(xlToLeft).Column
        .Cells(Counter_, col + 1) = "[" & timeStampObj.GetTimeStamp & "] " & contents
    End With
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：Getプロパティ関数
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
' 機能　　　　　　：Letプロパティ関数
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Let SheetName(value As String)
    Dim num As Long
    Dim targetFile As File

    SheetName_ = value

    With ThisWorkbook.Worksheets(SheetName_)
        MaxRowNumber_ = .Cells(.Rows.Count, 1).End(xlUp).row
        For num = 1 To MaxRowNumber_
            ' ファイルが存在する？
            If helperObj.IsFileExist(.Cells(num, "A")) Then
                ' B列が空（未処理）？
                If .Cells(num, "B") = "" Then
                    
                    ' 変換可能ファイルとしてカウントする
                    NumOfConvertibleFiles_ = NumOfConvertibleFiles_ + 1
                    
                    ' 当該行のB列以降の内容を全て消す
                    .Range(.Cells(num, "B"), .Cells(num, .Columns.Count)).Clear
                End If
            End If
        Next num
    End With

End Property
