Attribute VB_Name = "modPub"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public strAppPath As String 'Ӧ�ó���Ŀ¼
Public proName As String

Sub main()
    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    
    proName = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    frmMain.Show
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ����ļ����൱��˫���򿪶�Ӧ���ļ�
'��������openFileDbl
'��ڲ�����fileN,Ҫ�򿪵��ļ���ȫ·��
'��ע��sysdzw �� 20:55 2007-08-24 �ṩ
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub openFileDbl(hwnd As Long, fileN As String)
    ShellExecute hwnd, vbNullString, fileN, vbNullString, vbNullString, 1
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ����������ļ���������ֱ��д�ļ�
'��������writeToFile
'��ڲ���(����)��
'  strFileName �������ļ�����
'  strContent Ҫ���뵽�����ļ����ַ���
'  isCover �Ƿ񸲸Ǹ��ļ���Ĭ��Ϊ����
'����ֵ��True��False���ɹ��򷵻�ǰ�ߣ����򷵻غ���
'��ע��sysdzw �� 2007-5-2 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function writeToFile(ByVal strFileName$, ByVal strContent$, Optional isCover As Boolean = True) As Boolean
    On Error GoTo Err1
    Dim fileHandl%
    fileHandl = FreeFile
    If isCover Then
        Open strFileName For Output As #fileHandl
    Else
        Open strFileName For Append As #fileHandl
    End If
    Print #fileHandl, strContent
    Close #fileHandl
    writeToFile = True
    Exit Function
Err1:
    writeToFile = False
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'���ܣ������������ļ��������ļ�������
'��������fileStr
'��ڲ���(����)��
'  strFileName �������ļ�����
'����ֵ���ļ�������
'��ע��sysdzw �� 2007-5-3 �ṩ
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function fileStr(ByVal strFileName As String) As String
    On Error GoTo Err1
    Dim lFile&
    lFile = FreeFile
    Open strFileName For Input As #lFile
    fileStr = StrConv(InputB$(LOF(lFile), #lFile), vbUnicode)
    Close #lFile
    If Right(fileStr, 2) = vbCrLf Then fileStr = Left(fileStr, Len(fileStr) - 2)
    Exit Function
Err1:
'    MsgBox "�����ڸ��ļ�����ļ����ܷ��ʣ�" & vbCrLf & strFileName, vbExclamation
End Function
