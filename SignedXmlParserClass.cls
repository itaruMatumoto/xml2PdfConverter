VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SignedXmlParserClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' プロパティ変数
Private FileCollect_ As Collection          ' 公文書（通知書）
Private SocialInsurance_ As Boolean         ' =True：社会保険／=False：労働保険
Private XmlFileExist_ As Boolean            ' =True：XML形式の公文書がある／=False：無い
Private ArrivalNumber_ As String            ' 到達番号

' プライベート変数
Private xmlObj As MSXML2.DOMDocument60
Private helperObj As HelperClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス生成時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()

    Set FileCollect_ = New Collection
    Set helperObj = New HelperClass
    Set xmlObj = New MSXML2.DOMDocument60
    xmlObj.async = False

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：インスタンス破棄時
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set FileCollect_ = Nothing
    Set helperObj = Nothing
    Set xmlObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：署名付きXMLを解析する
' 第１引数（入力）：path As String          署名付きXMLのファイルパス
' 第２引数（出力）：rc As Long              リターンコード =0:正常
'                                                          =1:xmlファイルがない
'                                                          =2:xml解析エラーが発生した
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ParseXml(path As String, rc As Long)
    Dim eachNode As IXMLDOMNode
    Dim childNode As IXMLDOMNode
    Dim docName As String

    On Error GoTo ERROR_LABEL

    ' プライベート変数を初期化する
    InitializePrivateVariables

    rc = 0

    ' xmlが存在する？
    If helperObj.IsFileExist(path) Then
        ' xmlを取り込む
        xmlObj.Load (path)
        
        ' 到達番号を取得する
        Set eachNode = xmlObj.SelectSingleNode("//BODY/DOCNO")
        
        ArrivalNumber_ = eachNode.ChildNodes(0).Text
         
        ' 発出者を取得する
        Set eachNode = xmlObj.SelectSingleNode("//BODY/AUTHOR/AFF")
        
        If eachNode.ChildNodes(0).Text = "日本年金機構" Then
            SocialInsurance_ = True
        Else
            SocialInsurance_ = False
        End If

        ' 公文書リストを取得する
        Set eachNode = xmlObj.SelectSingleNode("//BODY/APPENDIX")
        
        ' 深さ優先探索を行い、公文書名を取得する
        Do
            ' 子要素を探索する
            For Each childNode In eachNode.ChildNodes
                Select Case childNode.nodeName
                    Case "DOCLINK"
                        docName = childNode.Attributes.getNamedItem("REF").NodeValue
                        Select Case docName
                            Case "henrei.xml"
                                ' 何もしない
                            Case Else
                                ' 拡張子がxml？
                                If Right(docName, 4) = ".xml" Then
                                    XmlFileExist_ = True
                                End If
                                FileCollect_.Add docName
                        End Select
                    Case Else
                        ' 何もしない
                End Select
            Next childNode
            ' 同じ階層の次のノードに移る
            Set eachNode = eachNode.NextSibling
        Loop While Not eachNode Is Nothing
    Else
        rc = 1
    End If
    
    GoTo FIN_LABEL

ERROR_LABEL:
    rc = 2

FIN_LABEL:

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：初期化
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub InitializePrivateVariables()

    Set FileCollect_ = New Collection
    XmlFileExist_ = False
    SocialInsurance_ = False
    ArrivalNumber_ = ""

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' 機能　　　　　　：Getプロパティ関数
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Property Get FileCollect() As Collection

    Set FileCollect = FileCollect_

End Property

Public Property Get SocialInsurance() As Boolean

    SocialInsurance = SocialInsurance_

End Property

Public Property Get XmlFileExist() As Boolean

    XmlFileExist = XmlFileExist_

End Property

Public Property Get ArrivalNumber() As String

    ArrivalNumber = ArrivalNumber_

End Property

