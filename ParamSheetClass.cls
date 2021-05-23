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

' プロパティ変数宣言
Private ExternalToolPath_ As String     ' 外部ツールパス
Private DebugOn_ As Boolean
Private Version_ As String

' プライベート変数宣言
Private helperObj As HelperClass

' 定数宣言
Const SheetName As String = "パラメタ"

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    Dim rangeObj As Variant
    Dim num As Long

    Set helperObj = New HelperClass
    DebugOn_ = False

    rangeObj = helperObj.GetRangeVariant(ThisWorkbook.path, SheetName)
    
    For num = LBound(rangeObj) To UBound(rangeObj)
        Select Case rangeObj(num, 1)
            Case "外部ツールパス"
                ExternalToolPath_ = rangeObj(num, 2)
            Case "プログラムバージョン"
                Version_ = rangeObj(num, 2)
            Case "debug"
                DebugOn_ = True
            Case Else
                ' その他
        End Select
    Next num
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：プロパティ関数
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
