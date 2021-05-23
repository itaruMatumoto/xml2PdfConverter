Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub Sleep Lib "kernel32" (ByVal ms As Long)

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
    ByVal hwndParent As Long, _
    ByVal hwndChildAfter As Long, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As Long

Private Declare Function SendMessageAny Lib "user32.dll" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As String) As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Const WM_SETTEXT = &HC
Private Const WM_KEYDOWN = &H100
Private Const VK_RETURN = &HD

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：メインコントローラー
' 第１引数（入力）：control As IRibbonControl　ユーザが操作したリボンの結果
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub MAIN_ON_ACTION(control As IRibbonControl)
    Dim paramObj As New ParamSheetClass

    'マクロ実行中の画面更新をしない
    Application.ScreenUpdating = False
    'マクロの実行中、ユーザーに入力を促すメッセージや警告メッセージを表示を抑止する
    Application.DisplayAlerts = False
    ' 数式再計算をマニュアルにする
    Application.Calculation = xlCalculationManual

    ' 各機能の実行を判定する
    Select Case control.ID
        Case "Button1"

            ' PDF変換メインコントローラを呼び出す
            Call Xml2PdfConvertMainController

        Case "Button9"
            ' バージョンを表示する
            MsgBox "プログラムバージョンは" & paramObj.Version & "です。", vbInformation, ThisWorkbook.name
        
        Case Else
    End Select
    
    Set paramObj = Nothing

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.StatusBar = ""
    Application.Calculation = xlCalculationAutomatic

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：PDF変換メインコントローラ
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub Xml2PdfConvertMainController()
    Dim rc As Long

    rc = MsgBox("新規に変換しますか？", vbYesNo + vbQuestion)

    ' 新規変換？
    If rc = vbYes Then
        With Application.FileDialog(msoFileDialogFolderPicker)
            If .Show = True Then
                ' ルートフォルダー以下を探索して、ファイルパスを取得する
                Call ExploreFolder(.SelectedItems.Item(1))
            Else
                MsgBox "キャンセルされました。", vbInformation
                GoTo FIN_LABEL
            End If
        End With
    End If

    ' XML->PDFに変換する
    Call ConvertTargetPath

FIN_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ルートフォルダー以下を探索して、ファイルパスを取得する
' 第１引数（入力）：path As String  ルートフォルダパス
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ExploreFolder(path As String)
    Dim exObj As New ExploreClass
    Dim helperObj As New HelperClass

    ' 条件を設定する
    exObj.SheetName = "xmlリスト"
    exObj.SearchPattern = "^[0-9]{17}[0].xml"      ' 鑑文書（署名付きxml）の正規表現

    ' フォルダを探索し、署名付きxmlを探す
    Call exObj.ExploreFolder(helperObj.GetFolderFso(path))

    Set exObj = Nothing
    Set helperObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：全パスをXML->PDFに変換する
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ConvertTargetPath()
    Dim shObj As New XmlListSheetClass

    shObj.SheetName = "xmlリスト"

    ' 変換対象ファイルあり？
    If shObj.NumOfConvertibleFiles > 0 Then
        ' 変換フォームを立ち上げる
        Call ShowPdfConvertForm
    Else
        MsgBox "PDF変換できるファイルがありません。", vbExclamation
    End If
    
    Set shObj = Nothing

End Sub


'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：変換フォームを表示する
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub ShowPdfConvertForm()

    Load PdfConvertForm
    PdfConvertForm.Show vbModeless

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：アクティブウインドウが印刷保存画面であるかを監視する
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub PollingPrintWindow()
    Dim hWnd As Long
    Dim titles As String * 1000
    Static cnt As Long
    Dim message As String

    Dim debugObj As DebugSheetClass

    Set debugObj = New DebugSheetClass

    Call debugObj.OutPut("ポーリングしています...")

    ' ウインドウハンドルを取得する
    hWnd = FindWindow(vbNullString, "印刷結果を名前を付けて保存")
    
    ' ウインドウハンドルが取得できた？
    If hWnd > 0 Then
        cnt = 0

        Dim hChildWnd As Long

        ' ファイル名のウインドウハンドルを求める
        hChildWnd = FindWindowEx(hWnd, 0, "DUIViewWndClassName", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "DirectUIHWND", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "FloatNotifySink", vbNullString)
        hChildWnd = FindWindowEx(hChildWnd, 0, "ComboBox", vbNullString)

        ' ファイル名のウインドウハンドルに対して、PDFファイルパスを送る
        Call SendMessageAny(hChildWnd, WM_SETTEXT, 0, PdfConvertForm.targetPdfPath)

        ' 保存(&S)のウインドウハンドルを求める
        hChildWnd = FindWindowEx(hWnd, 0, "Button", "保存(&S)")

        ' 保存(&S)のウインドウハンドルに対して、リターンキーを送る
        Call PostMessage(hChildWnd, WM_KEYDOWN, VK_RETURN, 0)

    Else
        cnt = cnt + 1

        message = "PDFに変換しています..." & PdfConvertForm.Progress & " - 監視中（" & cnt & "回）"

        Call PdfConvertForm.SetMessageTextBox(message)

        Application.OnTime Now + TimeSerial(0, 0, 5), "PollingPrintWindow"
    End If

    Set debugObj = Nothing

End Sub
