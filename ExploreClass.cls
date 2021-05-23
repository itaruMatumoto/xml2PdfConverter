VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExploreClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' プロパティ変数
Private SheetName_ As String            ' 探索結果を記録するシート名
Private SearchPattern_ As String        ' 探索するファイルパターン

' プライベート変数
Private shObj As XmlListSheetClass
Private xmlObj As SignedXmlParserClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set shObj = New XmlListSheetClass
    Set xmlObj = New SignedXmlParserClass

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Application.StatusBar = ""

    Set shObj = Nothing
    Set xmlObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：フォルダを探索する
' 第１引数（入力）：base_folder As folder 基準フォルダ
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ExploreFolder(baseFolder As Folder)
    Dim eachFolder As Folder
    Dim eachFile As File
    Dim eachDoc As Variant
    Dim mc As MatchCollection
    Dim path As String
    Dim rc As Long

    Dim helperObj As New HelperClass

    Application.StatusBar = "フォルダを探索します..." & baseFolder.path

    ' サブフォルダがある？
    If baseFolder.SubFolders.Count > 0 Then
        ' 再帰的に下位階層を探索する
        For Each eachFolder In baseFolder.SubFolders
            Call ExploreFolder(eachFolder)
        Next eachFolder
    End If
    
    ' 当該フォルダ直下のファイルを探す
    For Each eachFile In baseFolder.Files
        Set mc = helperObj.MatchRegExp(eachFile.name, SearchPattern_)
        
        ' パターンに該当した？
        If (mc.Count > 0) Then
            ' xmlを解析する
            Call xmlObj.ParseXml(eachFile.path, rc)
            
            ' 解析エラー？
            If rc > 0 Then
                MsgBox "xml解析エラーが発生しました。" & vbCrLf & eachFile.path, vbCritical
                Exit For
            End If
            
            ' 社会保険？
            If xmlObj.SocialInsurance Then
                ' 当該フォルダに存在する全ての公文書が存在するかチェックする
                For Each eachDoc In xmlObj.FileCollect
                    ' ファイルパスを取得する
                    path = eachFile.ParentFolder & "\" & eachDoc
                    ' ファイルパスが存在しない？
                    If Not helperObj.IsFileExist(path) Then
                        MsgBox "公文書が不足しています。不足分をダウンロードして下さい。" & vbCrLf & path, vbCritical
                        GoTo SKIP_LABEL
                    End If
                Next eachDoc
                    
                ' 公文書をxmlリストシートに書き出す
                For Each eachDoc In xmlObj.FileCollect
                    path = eachFile.ParentFolder & "\" & eachDoc
                    ' 当該フォルダにはxmlが存在する？
                    If xmlObj.XmlFileExist Then
                        If Right(eachDoc, 4) = ".xml" Then
                            ' xmlのみをシートに書き込む
                            Call shObj.Writer(path)
                        End If
                    Else
                        Call shObj.Writer(path)
                    End If
                Next eachDoc
            End If
            
            ' 当該フォルダのチェックが完了したので、次のフォルダーに移る
            Exit For
        End If
SKIP_LABEL:
    
    Next eachFile

    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：Letプロパティ関数
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Let SheetName(value As String)

    SheetName_ = value
    
    shObj.SheetName = SheetName_
    shObj.Clear

End Property

Public Property Let SearchPattern(value As String)

    SearchPattern_ = value

End Property
