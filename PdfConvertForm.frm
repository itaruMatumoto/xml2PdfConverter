VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PdfConvertForm 
   Caption         =   "PDF変換"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   10920
   Enabled         =   0   'False
   OleObjectBlob   =   "PdfConvertForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "PdfConvertForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function MessageBoxTimeoutA Lib "user32" ( _
        ByVal hWnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal uType As VbMsgBoxStyle, _
        ByVal wLanguageID As Long, _
        ByVal dwMilliseconds As Long) As Long

' プロパティ宣言
Private targetPdfPath_ As String
Private targetXmlPath_ As String

' プライベート変数
Private xmlListObj As XmlListSheetClass
Private shellObj As Object
Private helperObj As HelperClass
Private paramObj As ParamSheetClass
Private debugObj As DebugSheetClass
Private timeStampObj As TimeStampClass
Private mergerObj As MergerClass

Private defaultPrinterName As String

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ユーザフォーム生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub UserForm_Initialize()
    Dim list As Variant

    Set xmlListObj = New XmlListSheetClass
    Set shellObj = CreateObject("WScript.Shell")
    Set helperObj = New HelperClass
    Set paramObj = New ParamSheetClass
    Set debugObj = New DebugSheetClass
    Set timeStampObj = New TimeStampClass
    Set mergerObj = New MergerClass

    xmlListObj.SheetName = "xmlリスト"

    targetXmlPath_ = xmlListObj.Reader()

    Call debugObj.OutPut("ブラウザを表示します...")

    ' ブラウザを表示する
    With Me
        .WebBrowser1.Navigate targetXmlPath_
    End With
    
    ' PDF出力に切り替える
    list = Split(Application.ActivePrinter, " on ")
    defaultPrinterName = list(0)
    
    Call helperObj.ChangeActivePrinter("Microsoft Print to PDF")

FIN_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ユーザフォーム消滅時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub UserForm_Terminate()

    ' デフォルトプリンターに切り替える
    Call helperObj.ChangeActivePrinter(defaultPrinterName)

    Set xmlListObj = Nothing
    Set shellObj = Nothing
    Set helperObj = Nothing
    Set paramObj = Nothing
    Set debugObj = Nothing
    Set timeStampObj = Nothing
    Set mergerObj = Nothing

    Application.StatusBar = ""

    MsgBox "プログラムが終了しました。", vbInformation

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ドキュメント読み込み完了時に呼ばれるイベントハンドラー
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim rc As Boolean

    Call debugObj.OutPut("ドキュメントが読み込まれました...")

    With Me
        .MessageTextBox = "PDFに変換しています..." & Progress
    
        ' 変換するPDFパスを組み立てる
        targetPdfPath_ = helperObj.GetFileFso(targetXmlPath_).ParentFolder & "\temp_" & timeStampObj.GetTimeStamp2 & ".pdf"

        If helperObj.GetExtensionName(targetXmlPath_) = "xml" Then
        
            Call debugObj.OutPut("印刷を実行します...")

            ' 印刷実行
            ' 第1引数：OLECMDID_PRINT=6              　[ファイル]メニューの[印刷]
            ' 第2引数：OLECMDEXECOPT_DONTPROMPTUSER=2  ユーザーに入力を促すことなくコマンドを実行する
            .WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
    
        Else
            ' PDFファイルのコピーを作成する
            Call helperObj.CopyFile(targetXmlPath_, targetPdfPath_, rc)
        
            ' 成果物を作成する
            Call MakeDeliverables
        End If
    
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：PrintTemplateが生成された時に呼ばれるイベントハンドラー
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_PrintTemplateInstantiation(ByVal pDisp As Object)

    Call debugObj.OutPut("PrintTemplateが生成されました...")

    Application.OnTime Now + TimeSerial(0, 0, 5), "PollingPrintWindow"

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：PrintTemplateが消滅した時に呼ばれるイベントハンドラー
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub WebBrowser1_PrintTemplateTeardown(ByVal pDisp As Object)


    Call debugObj.OutPut("PrintTemplateが消滅しました...")
    Call debugObj.OutPut("PDFに変換しました...")

    Call xmlListObj.LogOnSheet("PDFに変換しました。")

    ' 成果物を作成する
    Call MakeDeliverables

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：成果物を作成する。
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub MakeDeliverables()
    Dim rc As Long

    ' 納品物を作成する
    Call mergerObj.MakeDeliverables(targetXmlPath_, targetPdfPath_, rc)
    
    ' 作成できた？
    Select Case rc
        Case Is = 1
            Call debugObj.OutPut("PDFをリネームしました...")
            Call xmlListObj.LogOnSheet("PDFをリネームしました。")
        Case Is = 2
            Call debugObj.OutPut("その他PDFをマージしました...")
            Call xmlListObj.LogOnSheet("その他PDFをマージしました。")
        Case Else
            ' 何もしない
    End Select

    ' シーケンスを続ける
    Call ContinueSequence(rc)

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シーケンスを続ける
' 第１引数（入力）：rc As Long          成果物の状態（>0：完成、=0：途中）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ContinueSequence(rc As Long)
    Dim m As String

    ' 次のxmlを読み込む
    targetXmlPath_ = xmlListObj.Reader()
    
    ' 次のxmlがある？
    If targetXmlPath_ <> "" Then
        ' 成果物が作成できた？
        If rc > 0 Then
            ' 一時停止を判断させる
            m = "中止しますか？" & vbCrLf & "指示が無ければ、３秒後に再開します。"
            rc = MessageBoxTimeoutA(0&, m, "確認", vbYesNo + vbQuestion + vbDefaultButton2, 0&, 3000)
            If rc = vbYes Then
                GoTo UNLOAD_LABEL
            End If
        End If

        Call debugObj.OutPut("ブラウザを表示します...")

        With Me
            ' ブラウザを表示する
            .WebBrowser1.Navigate targetXmlPath_
        End With
    
        ' シーケンスを継続する
        GoTo CONTINUE_LABEL
    
    End If
    
UNLOAD_LABEL:
    ' 次のxmlは無いため、フォームを閉じる（シーケンスを打ち切る）
    Unload Me

CONTINUE_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：メッセージを設定する
' 第１引数（入力）：message As String       メッセージ
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub SetMessageTextBox(message As String)

    With Me
        .MessageTextBox = message
    End With
    
End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：Getプロパティ関数
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Get targetPdfPath() As String

    targetPdfPath = targetPdfPath_

End Property

Public Property Get Progress() As String

    Progress = "(" & xmlListObj.Counter & "/" & xmlListObj.MaxRowNumber & "ファイル)"

End Property
