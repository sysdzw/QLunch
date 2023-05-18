Attribute VB_Name = "modPub"
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public strAppPath As String '应用程序目录
Public proName As String

Sub main()
    strAppPath = App.Path
    If Right(strAppPath, 1) <> "\" Then strAppPath = strAppPath & "\"
    
    proName = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    
    frmMain.Show
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：打开文件，相当于双击打开对应的文件
'函数名：openFileDbl
'入口参数：fileN,要打开的文件的全路径
'备注：sysdzw 于 20:55 2007-08-24 提供
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub openFileDbl(hwnd As Long, fileN As String)
    ShellExecute hwnd, vbNullString, fileN, vbNullString, vbNullString, 1
End Sub
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'功能：根据所给文件名和内容直接写文件
'函数名：writeToFile
'入口参数(如下)：
'  strFileName 所给的文件名；
'  strContent 要输入到上述文件的字符串
'  isCover 是否覆盖该文件，默认为覆盖
'返回值：True或False，成功则返回前者，否则返回后者
'备注：sysdzw 于 2007-5-2 提供
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
'功能：根据所给的文件名返回文件的内容
'函数名：fileStr
'入口参数(如下)：
'  strFileName 所给的文件名；
'返回值：文件的内容
'备注：sysdzw 于 2007-5-3 提供
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
'    MsgBox "不存在该文件或该文件不能访问！" & vbCrLf & strFileName, vbExclamation
End Function
