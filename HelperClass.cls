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
' �@�\�@�@�@�@�@�@�F�ΏۃV�[�g�̃Z���͈͑S����Variant�I�u�W�F�N�g�i�񎟌��z��j���擾����
'                   �񎟌��z��̓Y�����͂P�I���W���ł���
'
' ��P�����i���́j�Ftarget_sheet As String   �ΏۃV�[�g
' �Ԃ�l�@�@�@�@�@�FVariant�^�̓񎟌��z��
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
        
        ' ���o����V�[�g���W���O�z��ł��鎖��z�肵�āA�S�Ă̗�̂����̍ő�s�������߂�
        For num = 1 To colMax
            If .Cells(.Rows.Count, num).End(xlUp).row > rowMax Then
                rowMax = .Cells(.Rows.Count, num).End(xlUp).row
            End If
        Next num
        getRangeFromWorksheet = .Range(.Cells(1, 1), .Cells(rowMax, colMax))
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�F�w�肳�ꂽ�t�H���_�[�Ɉړ�����
' (����)  folderPath As String    �F�ړ��������t�H���_�[��΃p�X
' �߂�l�F��True�F�ړ������^��False�F�ړ��ł��Ȃ�����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ChangeFolder(folderPath As String)
    On Error GoTo label_error

    ChangeFolder = False

    '�J�����g�f�B���N�g����ύX
    If Left(folderPath, 2) = "\\" Then '�ŏ���2������\\�̏ꍇ�i�l�b�g���[�N�̏ꍇ�j
        'WSH�Ńf�B���N�g����ύX
        CreateObject("WScript.Shell").CurrentDirectory = folderPath
    Else '���[�J���h���C�u�̏ꍇ
        'ChDrive��ChDir�ŃJ�����g�h���C�u�ƃJ�����g�f�B���N�g����ύX
        Dim drive As String
        drive = Left(folderPath, 1)
        ChDrive drive
        ChDir folderPath
    End If

    ChangeFolder = True

    Exit Function

label_error:
    ' �G���[�������[�`���𖳌��ɂ���
    On Error GoTo 0
 
End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�Ώۃt�H���_�����݂��邩�ۂ���Ԃ�
' ��P�����i���́j�FfolderPath As String�@�Ώۃt�H���_�p�X
' �Ԃ�l�@�@�@�@�@�FTrue�i���݂���j�^False�i���݂��Ȃ��j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function IsFolderExist(folderPath As String) As Boolean

    IsFolderExist = False

    ' �t�H���_�����݂���H
    If Dir(folderPath, vbDirectory) <> "" Then
        IsFolderExist = True
    End If

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�Ώۃt�@�C�������݂��邩�ۂ���Ԃ�
' ��P�����i���́j�FfilePath As String�@�Ώۃt�@�C���p�X
' �Ԃ�l�@�@�@�@�@�FTrue�i���݂���j�^False�i���݂��Ȃ��j
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
' �@�\�@�@�@�@�@�@�F�Ώۃt�@�C���p�X�����l�[������
' ��P�����i���́j�FfilePath As String  �Ώۃt�@�C���p�X
' ��Q�����i���́j�Fname As String      ���l�[����̖��O
' ��R�����i�o�́j�Fresult as booelan   True�i�����j�^False�i���s�j
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
' �@�\�@�@�@�@�@�@�F�Ώۃt�@�C���p�X���폜����
' ��P�����i���́j�FfilePath As String  �Ώۃt�@�C���p�X
' ��Q�����i�o�́j�Fresult as booelan   True�i�����j�^False�i���s�j
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
' �@�\�@�@�@�@�@�@�F�Ώۃt�@�C���p�X���R�s�[����
' ��P�����i���́j�ForiginPath As String    �I���W�i���t�@�C���p�X
' ��Q�����i���́j�FcopyPath As String      �R�s�[��t�@�C���p�X
' ��R�����i�o�́j�Fresult as booelan       True�i�����j�^False�i���s�j
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
' �@�\�@�@�@�@�@�@�F�w��t�H���_�̐e�t�H���_��Ԃ�
' ��P�����i���́j�Ftarget_folder As String�@�Ώۃt�H���_�p�X
' �Ԃ�l�@�@�@�@�@�FString�^�̐e�t�H���_
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetParentFolderPath(targetFolder) As String

    GetParentFolderPath = Left(targetFolder, InStrRev(targetFolder, "\") - 1)

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�Ώۃt�H���_���̃T�u�t�H���_����Ԃ�
' ��P�����i���́j�FtargetFolder As String�@�Ώۃt�H���_�p�X
' �Ԃ�l�@�@�@�@�@�F�t�H���_����String�^�z��
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetSubFolder(targetFolder As String) As String()
    Dim buf As String
    Dim subFolderName As String
    Dim subFoldersList As String

    ' �Ώۃt�H���_������T��
    buf = Dir(targetFolder & "\*", vbDirectory)
    
    Do While buf <> ""
        Select Case buf
            Case ".", ".."
                ' �ǂݔ�΂�
            Case Else
                subFolderName = targetFolder & "\" & buf
                If GetAttr(subFolderName) And vbDirectory Then
                    ' �f���~�b�^�ivbCrLf�j��t�����ăT�u�t�H���_�������X�g������
                    subFoldersList = subFoldersList & buf & vbCrLf
                End If
        End Select

        buf = Dir()
    Loop
    
    ' �T�u�t�H���_��vbCrLf�ŕ������Ĕz��ɂ���
    Dim subFolderArray() As String
    
    subFolderArray = Split(subFoldersList, vbCrLf)
    
    'Call QuickSort(sub_folder_array, UBound(sub_folder_array), LBound(sub_folder_array))
    
    GetSubFolder = subFolderArray

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�z����N�C�b�N�\�[�g����
' ��P�����i���o�́j�FByRef data As Variant�@�\�[�g�O�z��^�\�[�g��z��
' ��Q�����i���́j�FByVal low As Long
' ��R�����i���́j�FByVal high As Long
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
' �@�\�@�@�@�@�@�@�F�V�[�g�̗L�����`�F�b�N����B
' ��P�����i���́@�FtargetBook As Workbook �`�F�b�N�ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String�@�`�F�b�N�ΏۂƂȂ�V�[�g��
' ��R�����i�o�́j�FFalse�F�V�[�g�����݂��Ȃ������ATrue�F�V�[�g�����݂���
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
' �@�\�@�@�@�@�@�@�F�V�[�g�̃t�B���^�[���������A�N���A��_���E�������̓��e���N���A����
' ��P�����i���́j�FtargetBook As Workbook �`�F�b�N�ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String�@�`�F�b�N�ΏۂƂȂ�V�[�g��
' ��R�����i���́j�FOptional base As String = "A1"  �N���A��_
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ClearSheet(targetBook As Workbook, targetSheet As String, Optional base As String = "A1")
    Dim autoFilterObj As autoFilter
    Dim rangeObj As Range

    With targetBook.Worksheets(targetSheet)
        ' �V�[�g�̃t�B���^�[����������
        Set autoFilterObj = .autoFilter
        If (Not autoFilterObj Is Nothing) Then
            If autoFilterObj.FilterMode Then
                .ShowAllData
            End If
        End If
        ' �N���A��_���E�������̓��e���N���A����
        Set rangeObj = .Range(.Range(base), .Cells(.Rows.Count, .Columns.Count))
        rangeObj.Clear
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�V�[�g�̃E�C���h�E�Œ��Ԃ���������
' ��P�����i���́j�FtargetBook As Workbook �`�F�b�N�ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String�@�`�F�b�N�ΏۂƂȂ�V�[�g��
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CancelFreezePanes(targetBook As Workbook, targetSheet As String)

    With targetBook.Worksheets(targetSheet)
        .Activate
        ActiveWindow.FreezePanes = False
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�V�[�g�𐶐�����
' ��P�����i���́j�FtargetBook As Workbook �����ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String�@�����ΏۂƂȂ�V�[�g��
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub CreateSheet(targetBook As Workbook, targetSheet As String)

    With targetBook
        .Worksheets.Add(After:=.Worksheets(.Worksheets.Count)).name = targetSheet
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�B���^�[��K�p����
' ��P�����i���́@�FtargetBook As Workbook      �`�F�b�N�ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String       �`�F�b�N�ΏۂƂȂ�V�[�g��
' ��R�����i���́j�Fposition As String          �t�B���^�[��������ʒu�i��J�������j
' ��S�����i���́j�Fvalue As String             �t�B���^�[�̑ΏۂƂȂ�l
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ApplyAutoFilter(targetBook As Workbook, targetSheet As String, position As String, value As String)

    With targetBook.Worksheets(targetSheet)
        .Range(position).autoFilter Field:=1, Criteria1:=value
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�B���^�[����������B
' ��P�����i���́@�FtargetBook As Workbook �`�F�b�N�ΏۂƂȂ郏�[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String�@�`�F�b�N�ΏۂƂȂ�V�[�g��
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
' �@�\�@�@�@�@�@�@�F��ԍ����p���ɕϊ�����
' ��P�����i���́@�FiCol As Long
' �Ԃ�l          �F�p��  As String
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
' �@�\�@�@�@�@�@�@�F�ΏۃV�[�g�̑Ώۃt�B�[���h�̃J�����ԍ���Ԃ�
' ��P�����i���́j�FsheetName as String
' ��Q�����i���́j�FcolumnName As String
' �Ԃ�l          �F�J�����ԍ� as Long
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
' �@�\�@�@�@�@�@�@�F�V�[�g���̃J����������J�����ԍ����擾����
' ��P�����i���́j�FsheetName As String         �V�[�g��
' ��Q�����i���́j�FColumnName As String        �J������
' ��R�����i�o�́j�FColumnNumber As Long        �J�����ԍ�
' ��S�����i�o�́j�FerrorMessages As String     �G���[���b�Z�[�W
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ConvColumnNumberFromKeyword(SheetName As String, columnName As String, columnNumber As Long, errorMessages As String)

    ' �J����������J�����ԍ����擾����
    columnNumber = getColumnNumber(SheetName, columnName)
    
    ' �J������������Ȃ��H
    If columnNumber = 0 Then
        errorMessages = errorMessages & vbCrLf & SheetName & ":" & columnName & "��������܂���B"
    End If

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�񎟌��z��̓��������o���A�ꎟ���z��ɕϊ�����B
' ��P�����i���́j�FtwoDimensionalArray As Variant�@�񎟌��z��i�P�I���W���j
' ��Q�����i���́j�FcolumnNumber As Long        �J�����ԍ��i�P�I���W���j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ConvTwoDimensionToOne(twoDimensionalArray As Variant, columnNumber As Long) As Variant
    Dim oneDimensionArray As Variant

    With Application.WorksheetFunction
        oneDimensionArray = .Transpose(twoDimensionalArray)
        ConvTwoDimensionToOne = .Index(oneDimensionArray, columnNumber)
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�񎟌��z�񂩂����s�����o���A�ꎟ���z��Ɋi�[����B
' ��P�����i���́j�FtwoDimensionalArray As Variant�@�񎟌��z��i�P�I���W���j
' ��Q�����i���́j�FrowNumber As Long        �s�ԍ��i�P�I���W���j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function ExtractOneRowFromTwoDimension(twoDimensionalArray As Variant, rowNumber As Long) As Variant

    With Application.WorksheetFunction
        ExtractOneRowFromTwoDimension = .Index(twoDimensionalArray, rowNumber)
    End With

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�R���N�V�����Ɉꎟ���z��̃A�C�e����ǉ�����B�z��͂P�I���W���ł���B
' ��P�����i���́j�FtargetCollection As Collection
' ��Q�����i���́j�FtargetArray As Variant
' ��R�����i���́j�FOptional offset As Long = 1     �ǉ����n�߂�z��̃I�t�Z�b�g
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
' �@�\�@�@�@�@�@�@�F�R���N�V�����̓��e���ꎟ���z��ɕϊ�����B�z��̓[���I���W���ł���B
' ��P�����i���́j�FtargetCollection As Collection
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
' �@�\�@�@�@�@�@�@�F�w�肳�ꂽ�N���������t�Ƃ��đÓ��ł��邩���`�F�b�N����
' ����            �FyearMonthDay As String      �N����
' �ԋp�l          �FTrue���Ó��ł���AFalse���Ó��ł͂Ȃ�
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
' �@�\�@�@�@�@�@�@�F���t������"yyyy/mm/dd"�`���ɕϊ�����
' ����            �FyearMonthDay As String      �N����
' �ԋp�l          �F�ϊ���̓��t�����iString�j
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
' �@�\�@�@�@�@�@�@�F���Y�Z���ɑ΂��āA�R�����g��ǋL����
' ����            �FtargetRange As Range    �Ώ�Range�I�u�W�F�N�g
'                 �Fcomment As String       �R�����g
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub AddComment(targetRange As Range, comment As String)

    Dim oldComment As String

    With targetRange.Cells(1, 1)
        '�Z���ɃR�����g���Ȃ��H
        If TypeName(.comment) = "Nothing" Then
            .AddComment comment
        Else
            '���ɃR�����g������ꍇ�͒ǋL����
            oldComment = .comment.Text
            .ClearComments
            .AddComment comment & vbCrLf & oldComment
        End If
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�w�肳�ꂽ�N�������P�ʂŉ����Z�����l��Ԃ�
' ��P�����i���́j�FyearMonth As String      �����N��
' ��Q�����i���́j�FaddOrSubMonth As Long    �����Z���錎��
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function CalculateYearMonth(YearMonth As String, addOrSubMonth As Long) As String
    ' �N
    Dim tmpYear As String
    ' ��
    Dim tmpMonth As String
    ' �N����
    Dim tmpYearMonthDay As Date

    ' �N�������Z����
    tmpYear = Left(YearMonth, 4)
    tmpMonth = Right(YearMonth, 2)
    tmpYearMonthDay = CDate(tmpYear + "/" + tmpMonth + "/01")
    tmpYearMonthDay = DateAdd("m", addOrSubMonth, tmpYearMonthDay)

    CalculateYearMonth = Format(tmpYearMonthDay, "yyyyMM")

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�Ώۃu�b�N�̑Ώۃ��[�N�V�[�g�̓��e���w�肳�ꂽ����L�[�ɏ����E�~���Ƀ\�[�g����
' ��P�����i���́j�FtargetBook As Workbook    �F�Ώۃ��[�N�u�b�N
' ��Q�����i���́j�FtargetSheet As String     �F�Ώۃ��[�N�V�[�g
' ��R�����i���́j�ForderArray(n,m) As Variant �F�\�[�g�L�[�y�ѕ��я����Ǘ�����񎟌��z��
'                                               ex orderArray(0,0) = ("C", True)�̏ꍇ�AC��������\�[�g
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
        ' �\�[�g�t�B�[���h���N���A����
        sortFieldsObj.Clear
        ' �\�[�g�t�B�[���h�ɉ��߂ă\�[�g�ʒu�A���я�����ݒ肷��
        For num = LBound(orderArray) To UBound(orderArray)
            columnPostion = orderArray(num, 0)
            columnOrder = orderArray(num, 1)
            If columnOrder Then
                sortFieldsObj.Add key:=xmlListObj.Columns(columnPostion), SortOn:=xlSortOnValues, order:=xlAscending, DataOption:=xlSortTextAsNumbers
            Else
                sortFieldsObj.Add key:=xmlListObj.Columns(columnPostion), SortOn:=xlSortOnValues, order:=xlDescending, DataOption:=xlSortTextAsNumbers
            End If
        Next num
        ' �\�[�g�͈͂�ݒ肷��
        Set sortRange = .UsedRange
    End With
    
    ' �\�[�g����
    With sortObj
        .SetRange sortRange
        .Header = xlYes     ' �w�b�_�����o���Ƃ��Ďg��
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlStroke
        .Apply             ' �\�[�g����
        .sortFields.Clear  ' �\�[�g�t�B�[���h���N���A����
    End With

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�������X�g�Ɋ܂܂�镶����啶���ɕϊ�����
' ��P�����i���́j�Fpattern as String�@�ϊ��O�̕�����
' ��Q�����i���́j�FOptional charList As String = "*"�@�������X�g
'                   (example) "[0-9]","[A-z]", "[0-9] | [A-z]"
' �Ԃ�l          �F�ϊ���̕�����  As String
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
' �@�\�@�@�@�@�@�@�F�������X�g�Ɋ܂܂�镶�����������ɕϊ�����
' ��P�����i���́j�Fpattern as String�@�ϊ��O�̕�����
' ��Q�����i���́j�FOptional charList As String = "*"�@�������X�g
'                   (example) "[�O-�X]","[�`-��]", "[�O-�X] | [�`-��]"
' �Ԃ�l          �F�ϊ���̕�����  As String
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
' �@�\�@�@�@�@�@�@�F�t�@�C���̊g���q��Ԃ�
' ��P�����i���́j�Fpath As String      �t�@�C���p�X
' �Ԃ�l          �F�t�@�C���̊g���q  As String
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetExtensionName(path As String) As String

    Dim fso As New Scripting.FileSystemObject
    
    GetExtensionName = fso.GetExtensionName(path)

    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F���e�ɑ΂��āA���K�\���ɍ��v���Ă���p�^�[����Ԃ�
' ��P�����i���́j�Fcontents As String      ���e
' ��Q�����i���́j�Fpattern As String       ���K�\��
' �Ԃ�l          �F���K�\���ɍ��v���Ă���p�^�[���Q    As MatchCollection
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
' �@�\�@�@�@�@�@�@�F�t�H���_�I�u�W�F�N�g�iFileSystemObject�j��Ԃ�
' ��P�����i���́j�Fpath As String          �t�H���_�p�X
' �Ԃ�l          �F�t�H���_�I�u�W�F�N�g    As folder
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetFolderFso(path As String) As Folder

    Dim fso As New Scripting.FileSystemObject
    
    Set GetFolderFso = fso.GetFolder(path)
    
    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�t�@�C���I�u�W�F�N�g�iFileSystemObject�j��Ԃ�
' ��P�����i���́j�Fpath As String          �t�@�C���p�X
' �Ԃ�l          �F�t�@�C���I�u�W�F�N�g    As file
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetFileFso(path As String) As File

    Dim fso As New Scripting.FileSystemObject
    
    Set GetFileFso = fso.GetFile(path)
    
    Set fso = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�|�[�g�ԍ��t���v�����^�����擾����
' ��P�����i���́j�FprinterName As String   �v�����^�[���iex."Microsoft Print to PDF"�j
' �Ԃ�l          �F�|�[�g�ԍ��t���v�����^�[���iex."Microsoft Print to PDF on xx"�j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Function GetPrinterNameWithPortNumber(printerName As String) As String
    Dim wsObj As Object, port As String
    
    Set wsObj = CreateObject("Wscript.Shell")
        
    port = wsObj.RegRead("HKCU\Software\Microsoft\Windows NT\CurrentVersion\Devices\" & printerName)

    GetPrinterNameWithPortNumber = printerName & " on " & Split(port, ",")(1)
    
    Set wsObj = Nothing

End Function

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�A�N�e�B�u�v�����^�[��؂�ւ���
' ��P�����i���́j�FprinterName As String   �|�[�g�ԍ��t���v�����^�[���iex."Microsoft Print to PDF on xx"�j
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ChangeActivePrinter(printerName As String)
    Dim wsObj As Object
    
    Set wsObj = CreateObject("Wscript.Network")
        
    wsObj.SetDefaultPrinter printerName
    
    Set wsObj = Nothing

End Sub