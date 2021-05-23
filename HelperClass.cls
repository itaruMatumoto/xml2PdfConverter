VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HelperClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象シートのセル範囲全部をVariantオブジェクト（二次元配列）を取得する
'                   二次元配列の添え字は１オリジンである
'
' 第１引数（入力）：target_sheet As String   対象シート
' 返り値　　　　　：Variant型の二次元配列
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetRangeVariant(targetBookPath As String, Optional targetSheet As String) As Variant

    Dim targetBook As Workbook

    If targetBookPath = ThisWorkbook.path Then
        With ThisWorkbook
            If targetSheet = "" Then
                GetRangeVariant = getRangeFromWorksheet(.Worksheets(1))
            Else
                GetRangeVariant = getRangeFromWorksheet(.Worksheets(targetSheet))
            End If
        End With
    Else
        Set targetBook = Application.Workbooks.Open(targetBookPath)
        With targetBook
            If targetSheet = "" Then
                GetRangeVariant = getRangeFromWorksheet(.Worksheets(1))
            Else
                GetRangeVariant = getRangeFromWorksheet(.Worksheets(targetSheet))
            End If
        End With
        targetBook.Close savechanges:=False
    End If

End Function

Private Function getRangeFromWorksheet(ws As Worksheet) As Variant
    Dim rowMax As Long
    Dim colMax As Long
    Dim num As Long

    With ws
        'colMax = .Cells(1, .Columns.Count).End(xlToLeft).Column
        colMax = ws.UsedRange.Columns.Count
        
        ' 抽出するシートがジャグ配列である事を想定して、全ての列のうちの最大行数を求める
        For num = 1 To colMax
            If .Cells(.Rows.Count, num).End(xlUp).row > rowMax Then
                rowMax = .Cells(.Rows.Count, num).End(xlUp).row
            End If
        Next num
        getRangeFromWorksheet = .Range(.Cells(1, 1), .Cells(rowMax, colMax))
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能：指定されたフォルダーに移動する
' (入力)  folderPath As String    ：移動したいフォルダー絶対パス
' 戻り値：＝True：移動した／＝False：移動できなかった
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ChangeFolder(folderPath As String)
    On Error GoTo label_error

    ChangeFolder = False

    'カレントディレクトリを変更
    If Left(folderPath, 2) = "\\" Then '最初の2文字が\\の場合（ネットワークの場合）
        'WSHでディレクトリを変更
        CreateObject("WScript.Shell").CurrentDirectory = folderPath
    Else 'ローカルドライブの場合
        'ChDriveとChDirでカレントドライブとカレントディレクトリを変更
        Dim drive As String
        drive = Left(folderPath, 1)
        ChDrive drive
        ChDir folderPath
    End If

    ChangeFolder = True

    Exit Function

label_error:
    ' エラー処理ルーチンを無効にする
    On Error GoTo 0
 
End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象フォルダが存在するか否かを返す
' 第１引数（入力）：folderPath As String　対象フォルダパス
' 返り値　　　　　：True（存在する）／False（存在しない）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function IsFolderExist(folderPath As String) As Boolean

    IsFolderExist = False

    ' フォルダが存在する？
    If Dir(folderPath, vbDirectory) <> "" Then
        IsFolderExist = True
    End If

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象ファイルが存在するか否かを返す
' 第１引数（入力）：filePath As String　対象ファイルパス
' 返り値　　　　　：True（存在する）／False（存在しない）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function IsFileExist(filePath As String) As Boolean
    
    Dim fsoObj As New FileSystemObject

    IsFileExist = False
    
    If fsoObj.FileExists(filePath) Then
        IsFileExist = True
    End If
        
    Set fsoObj = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象ファイルパスをリネームする
' 第１引数（入力）：filePath As String  対象ファイルパス
' 第２引数（入力）：name As String      リネーム後の名前
' 第３引数（出力）：result as booelan   True（成功）／False（失敗）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub RenameFile(filePath As String, name As String, result As Boolean)
    Dim fso As New Scripting.FileSystemObject
    Dim f As File

    On Error GoTo ERROR_LABEL
    
    result = False
    Set f = fso.GetFile(filePath)
    f.name = name
    result = True
    
ERROR_LABEL:
    
    Set fso = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象ファイルパスを削除する
' 第１引数（入力）：filePath As String  対象ファイルパス
' 第２引数（出力）：result as booelan   True（成功）／False（失敗）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub DeleteFile(filePath As String, result As Boolean)
    Dim fso As New Scripting.FileSystemObject
    Dim f As File

    On Error GoTo ERROR_LABEL
    
    result = False
    Set f = fso.GetFile(filePath)
    f.Delete
    result = True
    
ERROR_LABEL:
    
    Set fso = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象ファイルパスをコピーする
' 第１引数（入力）：originPath As String    オリジナルファイルパス
' 第２引数（入力）：copyPath As String      コピー先ファイルパス
' 第３引数（出力）：result as booelan       True（成功）／False（失敗）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CopyFile(originPath As String, copyPath As String, result As Boolean)
    Dim fso As New Scripting.FileSystemObject
    Dim f As File

    On Error GoTo ERROR_LABEL
    
    result = False
    Set f = fso.GetFile(originPath)
    f.Copy copyPath
    result = True
    
ERROR_LABEL:
    
    Set fso = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：指定フォルダの親フォルダを返す
' 第１引数（入力）：target_folder As String　対象フォルダパス
' 返り値　　　　　：String型の親フォルダ
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetParentFolderPath(targetFolder) As String

    GetParentFolderPath = Left(targetFolder, InStrRev(targetFolder, "\") - 1)

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象フォルダ内のサブフォルダ名を返す
' 第１引数（入力）：targetFolder As String　対象フォルダパス
' 返り値　　　　　：フォルダ名のString型配列
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetSubFolder(targetFolder As String) As String()
    Dim buf As String
    Dim subFolderName As String
    Dim subFoldersList As String

    ' 対象フォルダ直下を探す
    buf = Dir(targetFolder & "\*", vbDirectory)
    
    Do While buf <> ""
        Select Case buf
            Case ".", ".."
                ' 読み飛ばす
            Case Else
                subFolderName = targetFolder & "\" & buf
                If GetAttr(subFolderName) And vbDirectory Then
                    ' デリミッタ（vbCrLf）を付加してサブフォルダ名をリスト化する
                    subFoldersList = subFoldersList & buf & vbCrLf
                End If
        End Select

        buf = Dir()
    Loop
    
    ' サブフォルダをvbCrLfで分解して配列にする
    Dim subFolderArray() As String
    
    subFolderArray = Split(subFoldersList, vbCrLf)
    
    'Call QuickSort(sub_folder_array, UBound(sub_folder_array), LBound(sub_folder_array))
    
    GetSubFolder = subFolderArray

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：配列をクイックソートする
' 第１引数（入出力）：ByRef data As Variant　ソート前配列／ソート後配列
' 第２引数（入力）：ByVal low As Long
' 第３引数（入力）：ByVal high As Long
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub QuickSort(ByRef data As Variant, ByVal low As Long, ByVal high As Long)

    Dim l As Long
    Dim r As Long
    l = low
    r = high

    Dim pivot As Variant
    pivot = data((low + high) \ 2)

    Dim temp As Variant
    
    Do While (l <= r)
        Do While (data(l) < pivot And l < high)
            l = l + 1
        Loop
        Do While (pivot < data(r) And r > low)
            r = r - 1
        Loop
    
        If (l <= r) Then
            temp = data(l)
            data(l) = data(r)
            data(r) = temp
            l = l + 1
            r = r - 1
        End If
    Loop
    
    If (low < r) Then
        Call QuickSort(data, low, r)
    End If
    If (l < high) Then
        Call QuickSort(data, l, high)
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シートの有無をチェックする。
' 第１引数（入力　：targetBook As Workbook チェック対象となるワークブック
' 第２引数（入力）：targetSheet As String　チェック対象となるシート名
' 第３引数（出力）：False：シートが存在しなかった、True：シートが存在した
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function IsSheetExist(targetBook As Workbook, targetSheet As String) As Boolean
    Dim dummy As Variant
    
    On Error GoTo errorLabel
    
    IsSheetExist = True
    
    dummy = targetBook.Worksheets(targetSheet).Cells(1, "A")

    GoTo finLabel
    
errorLabel:
    IsSheetExist = False

finLabel:

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シートのフィルターを解除し、クリア基点より右下方向の内容をクリアする
' 第１引数（入力）：targetBook As Workbook チェック対象となるワークブック
' 第２引数（入力）：targetSheet As String　チェック対象となるシート名
' 第３引数（入力）：Optional base As String = "A1"  クリア基点
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ClearSheet(targetBook As Workbook, targetSheet As String, Optional base As String = "A1")
    Dim autoFilterObj As autoFilter
    Dim rangeObj As Range

    With targetBook.Worksheets(targetSheet)
        ' シートのフィルターを解除する
        Set autoFilterObj = .autoFilter
        If (Not autoFilterObj Is Nothing) Then
            If autoFilterObj.FilterMode Then
                .ShowAllData
            End If
        End If
        ' クリア基点より右下方向の内容をクリアする
        Set rangeObj = .Range(.Range(base), .Cells(.Rows.Count, .Columns.Count))
        rangeObj.Clear
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シートのウインドウ固定状態を解除する
' 第１引数（入力）：targetBook As Workbook チェック対象となるワークブック
' 第２引数（入力）：targetSheet As String　チェック対象となるシート名
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CancelFreezePanes(targetBook As Workbook, targetSheet As String)

    With targetBook.Worksheets(targetSheet)
        .Activate
        ActiveWindow.FreezePanes = False
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シートを生成する
' 第１引数（入力）：targetBook As Workbook 生成対象となるワークブック
' 第２引数（入力）：targetSheet As String　生成対象となるシート名
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CreateSheet(targetBook As Workbook, targetSheet As String)

    With targetBook
        .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).name = targetSheet
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：フィルターを適用する
' 第１引数（入力　：targetBook As Workbook      チェック対象となるワークブック
' 第２引数（入力）：targetSheet As String       チェック対象となるシート名
' 第３引数（入力）：position As String          フィルターをかける位置（列カラム名）
' 第４引数（入力）：value As String             フィルターの対象となる値
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ApplyAutoFilter(targetBook As Workbook, targetSheet As String, position As String, value As String)

    With targetBook.Worksheets(targetSheet)
        .Range(position).autoFilter Field:=1, Criteria1:=value
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：フィルターを解除する。
' 第１引数（入力　：targetBook As Workbook チェック対象となるワークブック
' 第２引数（入力）：targetSheet As String　チェック対象となるシート名
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CancelAutoFilter(targetBook As Workbook, targetSheet As String)
    Dim filterObj As autoFilter

    With targetBook.Worksheets(targetSheet)
        Set filterObj = .autoFilter
        If (Not filterObj Is Nothing) Then
            If filterObj.FilterMode Then
                .Range("A1").autoFilter
            End If
        End If
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：列番号を英字に変換する
' 第１引数（入力　：iCol As Long
' 返り値          ：英字  As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvertToLetter(iCol As Long) As String
    Dim a As Long
    Dim b As Long
    Dim c As Long

    a = iCol
    c = iCol
    
    ConvertToLetter = ""
    Do While c > 0
        a = Int((c - 1) / 26)
        b = (c - 1) Mod 26
        ConvertToLetter = Chr(b + 65) & ConvertToLetter
        c = a
    Loop

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象シートの対象フィールドのカラム番号を返す
' 第１引数（入力）：sheetName as String
' 第２引数（入力）：columnName As String
' 返り値          ：カラム番号 as Long
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function getColumnNumber(SheetName As String, columnName As String) As Long
    Dim columnNames As Variant
    Dim num As Long
    
    On Error GoTo SKIP_LABEL
    
    getColumnNumber = 0
    
    With ThisWorkbook.Worksheets(SheetName)
        columnNames = .Rows(1)
        For num = LBound(columnNames, 2) To UBound(columnNames, 2)
            If columnNames(1, num) Like columnName & "*" Then
                getColumnNumber = num
                GoTo SKIP_LABEL
            End If
        Next num
    End With

SKIP_LABEL:

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：シート内のカラム名からカラム番号を取得する
' 第１引数（入力）：sheetName As String         シート名
' 第２引数（入力）：ColumnName As String        カラム名
' 第３引数（出力）：ColumnNumber As Long        カラム番号
' 第４引数（出力）：errorMessages As String     エラーメッセージ
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ConvColumnNumberFromKeyword(SheetName As String, columnName As String, columnNumber As Long, errorMessages As String)

    ' カラム名からカラム番号を取得する
    columnNumber = getColumnNumber(SheetName, columnName)
    
    ' カラムが見つからない？
    If columnNumber = 0 Then
        errorMessages = errorMessages & vbCrLf & SheetName & ":" & columnName & "が見つかりません。"
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：二次元配列の特定列を取り出し、一次元配列に変換する。
' 第１引数（入力）：twoDimensionalArray As Variant　二次元配列（１オリジン）
' 第２引数（入力）：columnNumber As Long        カラム番号（１オリジン）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvTwoDimensionToOne(twoDimensionalArray As Variant, columnNumber As Long) As Variant
    Dim oneDimensionArray As Variant

    With Application.WorksheetFunction
        oneDimensionArray = .Transpose(twoDimensionalArray)
        ConvTwoDimensionToOne = .Index(oneDimensionArray, columnNumber)
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：二次元配列から特定行を取り出し、一次元配列に格納する。
' 第１引数（入力）：twoDimensionalArray As Variant　二次元配列（１オリジン）
' 第２引数（入力）：rowNumber As Long        行番号（１オリジン）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ExtractOneRowFromTwoDimension(twoDimensionalArray As Variant, rowNumber As Long) As Variant

    With Application.WorksheetFunction
        ExtractOneRowFromTwoDimension = .Index(twoDimensionalArray, rowNumber)
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：コレクションに一次元配列のアイテムを追加する。配列は１オリジンである。
' 第１引数（入力）：targetCollection As Collection
' 第２引数（入力）：targetArray As Variant
' 第３引数（入力）：Optional offset As Long = 1     追加を始める配列のオフセット
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub AddCollection(targetCollection As Collection, targetArray As Variant, Optional offset As Long = 1)
    Dim num As Long
    
    For num = LBound(targetArray) + offset - 1 To UBound(targetArray)
        If Not IsEmpty(targetArray) Then
            targetCollection.Add targetArray(num)
        End If
    Next num

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：コレクションの内容を一次元配列に変換する。配列はゼロオリジンである。
' 第１引数（入力）：targetCollection As Collection
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvCollectionToArray(targetCollection As Collection) As Variant
    Dim num As Long
    Dim targetArray As Variant
    
    ReDim targetArray(targetCollection.Count - 1)
    
    For num = 1 To targetCollection.Count
        targetArray(num - 1) = targetCollection(num)
    Next num

    ConvCollectionToArray = targetArray

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：指定された年月日が日付として妥当であるかをチェックする
' 入力            ：yearMonthDay As String      年月日
' 返却値          ：True＝妥当である、False＝妥当ではない
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function IsValidDate(yearMonthDay As Variant) As Boolean
    
    IsValidDate = False
    
    If IsDate(yearMonthDay) Then
        IsValidDate = True
    Else
        If IsNumeric(yearMonthDay) And Len(yearMonthDay) = 8 Then
            IsValidDate = IsDate(Format(yearMonthDay, "####/##/##"))
        End If
    End If

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：日付書式を"yyyy/mm/dd"形式に変換する
' 入力            ：yearMonthDay As String      年月日
' 返却値          ：変換後の日付書式（String）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ChangeFormatDate(yearMonthDay As Variant) As String

    ChangeFormatDate = ""

    If IsValidDate(yearMonthDay) Then
        If IsDate(yearMonthDay) Then
            ChangeFormatDate = yearMonthDay
        Else
            If IsNumeric(yearMonthDay) And Len(yearMonthDay) = 8 Then
                ChangeFormatDate = Format(yearMonthDay, "####/##/##")
            End If
        End If
    End If

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：当該セルに対して、コメントを追記する
' 入力            ：targetRange As Range    対象Rangeオブジェクト
'                 ：comment As String       コメント
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub AddComment(targetRange As Range, comment As String)

    Dim oldComment As String

    With targetRange.Cells(1, 1)
        'セルにコメントがない？
        If TypeName(.comment) = "Nothing" Then
            .AddComment comment
        Else
            '既にコメントがある場合は追記する
            oldComment = .comment.Text
            .ClearComments
            .AddComment comment & vbCrLf & oldComment
        End If
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：指定された年月を月単位で加減算した値を返す
' 第１引数（入力）：yearMonth As String      処理年月
' 第２引数（入力）：addOrSubMonth As Long    加減算する月数
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function CalculateYearMonth(YearMonth As String, addOrSubMonth As Long) As String
    ' 年
    Dim tmpYear As String
    ' 月
    Dim tmpMonth As String
    ' 年月日
    Dim tmpYearMonthDay As Date

    ' 年月を加算する
    tmpYear = Left(YearMonth, 4)
    tmpMonth = Right(YearMonth, 2)
    tmpYearMonthDay = CDate(tmpYear + "/" + tmpMonth + "/01")
    tmpYearMonthDay = DateAdd("m", addOrSubMonth, tmpYearMonthDay)

    CalculateYearMonth = Format(tmpYearMonthDay, "yyyyMM")

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：対象ブックの対象ワークシートの内容を指定された列をキーに昇順・降順にソートする
' 第１引数（入力）：targetBook As Workbook    ：対象ワークブック
' 第２引数（入力）：targetSheet As String     ：対象ワークシート
' 第３引数（入力）：orderArray(n,m) As Variant ：ソートキー及び並び順を管理する二次元配列
'                                               ex orderArray(0,0) = ("C", True)の場合、C列を昇順ソート
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub SortWorksheet(targetBook As Workbook, targetSheet As String, orderArray() As Variant)
    Dim num As Long
    Dim columnPostion As String
    Dim columnOrder As Boolean

    Dim xmlListObj As Worksheet
    Dim sortObj As sort
    Dim sortFieldsObj As sortFields
    Dim sortRange As Range

    Set xmlListObj = targetBook.Worksheets(targetSheet)
    Set sortObj = xmlListObj.sort
    Set sortFieldsObj = sortObj.sortFields

    With xmlListObj
        ' ソートフィールドをクリアする
        sortFieldsObj.Clear
        ' ソートフィールドに改めてソート位置、並び順等を設定する
        For num = LBound(orderArray) To UBound(orderArray)
            columnPostion = orderArray(num, 0)
            columnOrder = orderArray(num, 1)
            If columnOrder Then
                sortFieldsObj.Add key:=xmlListObj.Columns(columnPostion), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
            Else
                sortFieldsObj.Add key:=xmlListObj.Columns(columnPostion), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortTextAsNumbers
            End If
        Next num
        ' ソート範囲を設定する
        Set sortRange = .UsedRange
    End With
    
    ' ソートする
    With sortObj
        .SetRange sortRange
        .Header = xlYes     ' ヘッダを見出しとして使う
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlStroke
        .Apply             ' ソートする
        .sortFields.Clear  ' ソートフィールドをクリアする
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：文字リストに含まれる文字を大文字に変換する
' 第１引数（入力）：pattern as String　変換前の文字列
' 第２引数（入力）：Optional charList As String = "*"　文字リスト
'                   (example) "[0-9]","[A-z]", "[0-9] | [A-z]"
' 返り値          ：変換後の文字列  As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvertToFullWidth(pattern As String, Optional charList As String = "*") As String
    Dim num As Long
    Dim oneChar As String

    For num = 1 To Len(pattern)
        oneChar = Mid(pattern, num, 1)
        If oneChar Like charList Then
            ConvertToFullWidth = ConvertToFullWidth & StrConv(oneChar, vbWide)
        Else
            ConvertToFullWidth = ConvertToFullWidth & oneChar
        End If
    Next num

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：文字リストに含まれる文字を小文字に変換する
' 第１引数（入力）：pattern as String　変換前の文字列
' 第２引数（入力）：Optional charList As String = "*"　文字リスト
'                   (example) "[０-９]","[Ａ-ｚ]", "[０-９] | [Ａ-ｚ]"
' 返り値          ：変換後の文字列  As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvertToHalfSize(pattern As String, Optional charList As String = "*") As String
    Dim num As Long
    Dim oneChar As String

    For num = 1 To Len(pattern)
        oneChar = Mid(pattern, num, 1)
        If oneChar Like charList Then
            ConvertToHalfSize = ConvertToHalfSize & StrConv(oneChar, vbNarrow)
        Else
            ConvertToHalfSize = ConvertToHalfSize & oneChar
        End If
    Next num

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ファイルの拡張子を返す
' 第１引数（入力）：path As String      ファイルパス
' 返り値          ：ファイルの拡張子  As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetExtensionName(path As String) As String

    Dim fso As New Scripting.FileSystemObject
    
    GetExtensionName = fso.GetExtensionName(path)

    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：内容に対して、正規表現に合致しているパターンを返す
' 第１引数（入力）：contents As String      内容
' 第２引数（入力）：pattern As String       正規表現
' 返り値          ：正規表現に合致しているパターン群    As MatchCollection
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function MatchRegExp(contents As String, pattern As String) As MatchCollection
    Dim reObj As New regExp

    With reObj
        .pattern = pattern
        .Global = True
        .IgnoreCase = False
        Set MatchRegExp = .execute(contents)
    End With

    Set reObj = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：フォルダオブジェクト（FileSystemObject）を返す
' 第１引数（入力）：path As String          フォルダパス
' 返り値          ：フォルダオブジェクト    As folder
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetFolderFso(path As String) As Folder

    Dim fso As New Scripting.FileSystemObject
    
    Set GetFolderFso = fso.GetFolder(path)
    
    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ファイルオブジェクト（FileSystemObject）を返す
' 第１引数（入力）：path As String          ファイルパス
' 返り値          ：ファイルオブジェクト    As file
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetFileFso(path As String) As File

    Dim fso As New Scripting.FileSystemObject
    
    Set GetFileFso = fso.GetFile(path)
    
    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：ポート番号付きプリンタ名を取得する
' 第１引数（入力）：printerName As String   プリンター名（ex."Microsoft Print to PDF"）
' 返り値          ：ポート番号付きプリンター名（ex."Microsoft Print to PDF on xx"）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetPrinterNameWithPortNumber(printerName As String) As String
    Dim wsObj As Object, port As String
    
    Set wsObj = CreateObject("Wscript.Shell")
        
    port = wsObj.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices\" & printerName)

    GetPrinterNameWithPortNumber = printerName & " on " & Split(port, ",")(1)
    
    Set wsObj = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：アクティブプリンターを切り替える
' 第１引数（入力）：printerName As String   ポート番号付きプリンター名（ex."Microsoft Print to PDF on xx"）
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ChangeActivePrinter(printerName As String)
    Dim wsObj As Object
    
    Set wsObj = CreateObject("Wscript.Network")
        
    wsObj.SetDefaultPrinter printerName
    
    Set wsObj = Nothing

End Sub
