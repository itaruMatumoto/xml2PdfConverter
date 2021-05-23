VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MergerClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

' プライベート変数
Private externalToolPath As String
Private baseFolder As Folder
Private convertedFilesDict As Dictionary
Private signedXmlFile As File

Private helperObj As HelperClass
Private xmlObj As SignedXmlParserClass
Private shellObj As Object
Private debugObj As DebugSheetClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()
    
    Set convertedFilesDict = New Dictionary
    Set helperObj = New HelperClass
    Set xmlObj = New SignedXmlParserClass
    Set shellObj = CreateObject("WScript.Shell")
    Set debugObj = New DebugSheetClass

    externalToolPath = "qpdf.exe"

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set convertedFilesDict = Nothing
    Set helperObj = Nothing
    Set xmlObj = Nothing
    Set shellObj = Nothing
    Set debugObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：納品物を作成する
' 第１引数（入力）：xmlFilePath As String   xmlファイルパス
' 第２引数（入力）：pdfFilePath As String   変換済みPDFファイルパス
' 第３引数（出力）：rc As Long              リターンコード =0：未処理
'                                                          =1：リネームした
'                                                          =2：マージした
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub MakeDeliverables(xmlFilePath As String, pdfFilePath As String, rc As Long)

    rc = 0

    ' マージ環境を初期化する
    Call InitializeMergeEnvironment(xmlFilePath)

    ' xmlファイルパスをキー、変換済みPDFファイルパスを値として、辞書に登録する
    convertedFilesDict.Add xmlFilePath, pdfFilePath

    ' マージファイルを作成できるか確認する
    If IsPossibleMerge Then
        Select Case xmlObj.FileCollect.Count
            Case Is = 1
                ' 変換済みPDFファイルをリネームする
                Call RenamePdf(pdfFilePath)
                rc = 1
            Case Is > 1
                ' マージファイルを作成する
                Call MergeFiles
                ' 変換済みPDFを全て削除する
                Call RemoveConvertedFiles
                rc = 2
            Case Else
                ' 何もしない
        End Select
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：baseFolder未設定時、baseFolder変更時にマージ環境を初期化する
' 第１引数（入力）：xmlFilePath As String       xmlファイルパス
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub InitializeMergeEnvironment(xmlFilePath As String)
    Dim f As File

    Set f = helperObj.GetFileFso(xmlFilePath)

    ' baseFolder未設定？
    If baseFolder Is Nothing Then
        Set baseFolder = helperObj.GetFolderFso(f.ParentFolder)
        ' 署名付きxmlを解析する
        Call ParseSignedXml
    Else
        ' ベースフォルダが変わった？
        If baseFolder.path <> f.ParentFolder Then
            convertedFilesDict.RemoveAll
            Set baseFolder = helperObj.GetFolderFso(f.ParentFolder)
            ' 署名付きxmlを解析する
            Call ParseSignedXml
        End If
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ファイル名をリネームする
' 第１引数（入力）：pdfFilePath As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub RenamePdf(pdfFilePath As String)
    Dim result As Boolean
    Dim targetPdfPath As String
    Dim name As String

    Call debugObj.OutPut("リネームします...")

    targetPdfPath = Replace(signedXmlFile.path, ".xml", ".pdf")

    ' 変換済みPDFが既にある？
    If helperObj.IsFileExist(targetPdfPath) Then
        ' 変換済みPDFを削除する
        Call helperObj.DeleteFile(targetPdfPath, result)
    End If

    name = Replace(signedXmlFile.name, ".xml", ".pdf")

    ' 変換済みPDFファイルをリネームする
    Call helperObj.RenameFile(pdfFilePath, name, result)

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：変換済みPDFを削除する
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub RemoveConvertedFiles()
    Dim var As Variant
    Dim f As File
    Dim num As Long
    
    var = convertedFilesDict.Items
    
    For num = LBound(var) To UBound(var)
        Set f = helperObj.GetFileFso(CStr(var(num)))
        f.Delete
    Next num

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ファイルをマージする
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub MergeFiles()
    Dim command As String
    Dim pdf As String
    Dim f As File

    Call debugObj.OutPut("マージします...")

    ' マージ後のPDF名を決める
    pdf = Replace(signedXmlFile.path, ".xml", ".pdf")
    
    ' PDFがあれば削除する
    If helperObj.IsFileExist(pdf) Then
        Set f = helperObj.GetFileFso(pdf)
        f.Delete
    End If
       
    ' コマンドを組み立てる
    command = MakeCommand
    
    Call debugObj.OutPut(command)
    
    ' コマンドを実行する
    shellObj.Run command, 0, True
        
    ' マージ後のPDFが出来るまで待つ
    While (Not helperObj.IsFileExist(pdf))
        Sleep 1000
        DoEvents
    Wend

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：コマンドを組み立てる
' 注意　qpdfは、パス名に空白があると動かない。
'       コマンドを組み立てる際にパスをダブルクオーテーションで囲む必要がある。
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function MakeCommand() As String
    Dim pdfAfterMerging As String
    Dim pdfMergeList As String
    Dim eachValue As Variant
    Dim inputPath As String
    Dim eachPath As String

    ' マージ後のPDF名を決める
    pdfAfterMerging = EnclosePathDoubleQuotes(Replace(signedXmlFile.path, ".xml", ".pdf"))
    
    ' マージリストを作成する
    For Each eachValue In xmlObj.FileCollect
        ' パスに組み立てる
        eachPath = baseFolder & "\" & eachValue
        
        ' xmlファイル？
        If helperObj.GetExtensionName(eachPath) = "xml" Then
            If pdfMergeList = "" Then
                pdfMergeList = EnclosePathDoubleQuotes(convertedFilesDict.Item(eachPath))
            Else
                pdfMergeList = pdfMergeList & " " & EnclosePathDoubleQuotes(convertedFilesDict.Item(eachPath))
            End If
        Else
            If pdfMergeList = "" Then
                pdfMergeList = EnclosePathDoubleQuotes(eachPath)
            Else
                pdfMergeList = pdfMergeList & " " & EnclosePathDoubleQuotes(eachPath)
            End If
        End If
        
        ' １つ目のパスを入力パスとして覚えておく
        If inputPath = "" Then
            inputPath = pdfMergeList
        End If

    Next eachValue
    
    ' コマンドを組み立てる
    ' （注意）qpdfは、パス名に空白があると動かない。
    '         コマンドを組み立てる際にパスをダブルクオーテーションで囲む必要がある。
    MakeCommand = externalToolPath & " " & inputPath & " " & _
                    "--pages " & pdfMergeList & " " & _
                    "-- " & pdfAfterMerging

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：パスに空白があるケースを想定して、パスをダブルクオーテーションで囲む
' 第１引数（入力）：path As String      ファイルパス
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function EnclosePathDoubleQuotes(path As String) As String

    EnclosePathDoubleQuotes = """" & path & """"

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：鑑文書（署名付きxml）を解析する
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ParseSignedXml()
    Dim rc As Long
    
    ' 鑑文書を取得する
    Set signedXmlFile = GetSignedXml()
    
    ' 鑑文書を解析する
    Call xmlObj.ParseXml(signedXmlFile.path, rc)

    ' 解析エラー？
    If rc > 0 Then
        MsgBox "xml解析エラーが発生しました。" & vbCrLf & signedXmlFile.path, vbCritical
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：鑑文書（署名付きxml）を返す
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function GetSignedXml() As File
    Dim f As File
    Dim mc As MatchCollection
    
    For Each f In baseFolder.Files
        ' 署名付きxmlの正規表現パターンでマッチングする
        Set mc = helperObj.MatchRegExp(f.name, "^[0-9]{17}[0].xml")
        
        ' パターンあり？
        If (mc.Count > 0) Then
            Set GetSignedXml = f
            GoTo FIN_LABEL
        End If
    Next f

FIN_LABEL:

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：マージファイルを作成できるか否かを返す
' 返り値          ：As Boolean　=True:作成できる／=False:作成できない
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Function IsPossibleMerge() As Boolean
    Dim eachValue As Variant
    Dim path As String
    
    IsPossibleMerge = True
    
    For Each eachValue In xmlObj.FileCollect
        path = baseFolder & "\" & eachValue
        ' xmlファイル？
        If helperObj.GetExtensionName(path) = "xml" Then
            ' 変換済みPDFは存在しない？
            If Not convertedFilesDict.Exists(path) Then
                IsPossibleMerge = False
                ' １ファイルでもPDF未変換の場合、マージ出来ないので抜ける
                Exit For
            End If
        End If
    Next eachValue

End Function

