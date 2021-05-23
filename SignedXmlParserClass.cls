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

' �v���p�e�B�ϐ�
Private FileCollect_ As Collection          ' �������i�ʒm���j
Private SocialInsurance_ As Boolean         ' =True�F�Љ�ی��^=False�F�J���ی�
Private XmlFileExist_ As Boolean            ' =True�FXML�`���̌�����������^=False�F����
Private ArrivalNumber_ As String            ' ���B�ԍ�

' �v���C�x�[�g�ϐ�
Private xmlObj As MSXML2.DOMDocument60
Private helperObj As HelperClass

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Initialize()

    Set FileCollect_ = New Collection
    Set helperObj = New HelperClass
    Set xmlObj = New MSXML2.DOMDocument60
    xmlObj.async = False

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�C���X�^���X�j����
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub Class_Terminate()

    Set FileCollect_ = Nothing
    Set helperObj = Nothing
    Set xmlObj = Nothing

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�F�����t��XML����͂���
' ��P�����i���́j�Fpath As String          �����t��XML�̃t�@�C���p�X
' ��Q�����i�o�́j�Frc As Long              ���^�[���R�[�h =0:����
'                                                          =1:xml�t�@�C�����Ȃ�
'                                                          =2:xml��̓G���[����������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Public Sub ParseXml(path As String, rc As Long)
    Dim eachNode As IXMLDOMNode
    Dim childNode As IXMLDOMNode
    Dim docName As String

    On Error GoTo ERROR_LABEL

    ' �v���C�x�[�g�ϐ�������������
    InitializePrivateVariables

    rc = 0

    ' xml�����݂���H
    If helperObj.IsFileExist(path) Then
        ' xml����荞��
        xmlObj.Load (path)
        
        ' ���B�ԍ����擾����
        Set eachNode = xmlObj.SelectSingleNode("//BODY/DOCNO")
        
        ArrivalNumber_ = eachNode.ChildNodes(0).Text
         
        ' ���o�҂��擾����
        Set eachNode = xmlObj.SelectSingleNode("//BODY/AUTHOR/AFF")
        
        If eachNode.ChildNodes(0).Text = "���{�N���@�\" Then
            SocialInsurance_ = True
        Else
            SocialInsurance_ = False
        End If

        ' ���������X�g���擾����
        Set eachNode = xmlObj.SelectSingleNode("//BODY/APPENDIX")
        
        ' �[���D��T�����s���A�����������擾����
        Do
            ' �q�v�f��T������
            For Each childNode In eachNode.ChildNodes
                Select Case childNode.nodeName
                    Case "DOCLINK"
                        docName = childNode.Attributes.getNamedItem("REF").NodeValue
                        Select Case docName
                            Case "henrei.xml"
                                ' �������Ȃ�
                            Case Else
                                ' �g���q��xml�H
                                If Right(docName, 4) = ".xml" Then
                                    XmlFileExist_ = True
                                End If
                                FileCollect_.Add docName
                        End Select
                    Case Else
                        ' �������Ȃ�
                End Select
            Next childNode
            ' �����K�w�̎��̃m�[�h�Ɉڂ�
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
' �@�\�@�@�@�@�@�@�F������
'--------+---------+---------+---------+---------+---------+---------+---------+---------+
Private Sub InitializePrivateVariables()

    Set FileCollect_ = New Collection
    XmlFileExist_ = False
    SocialInsurance_ = False
    ArrivalNumber_ = ""

End Sub

'--------+---------+---------+---------+---------+---------+---------+---------+---------+
' �@�\�@�@�@�@�@�@�FGet�v���p�e�B�֐�
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
