VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "快速启动"
   ClientHeight    =   9075
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14520
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   9075
   ScaleWidth      =   14520
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "所在目录"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4440
      TabIndex        =   10
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   11040
      TabIndex        =   8
      Text            =   "移动"
      Top             =   4320
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "关于(&A)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13080
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "打开选中项"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   10920
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存列表"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5640
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "删除选中项"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   0
      ToolTipText     =   "双击打开"
      Top             =   840
      Width           =   10335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "移动"
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   11040
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "提示：要添加的文件或文件夹可拖拽到下面列表区"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   315
      Left            =   6840
      TabIndex        =   6
      Top             =   240
      Width           =   5280
   End
   Begin VB.Label lblTip 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "状态:"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   4
      Top             =   8040
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strSetFile As String
Dim intMoveStart%, intMoveEnd%
Dim isMove As Boolean

Private Sub Command1_Click()
    List2.RemoveItem (List1.ListIndex)
    List1.RemoveItem (List1.ListIndex)
    
    writeToFile strSetFile, getListAllItem(List2)
End Sub

Private Sub Command2_Click()
    writeToFile strSetFile, getListAllItem(List2)
    MsgBox "保存成功！", vbInformation
End Sub

Private Sub Command3_Click()
    Call List1_DblClick
End Sub

Private Sub Command4_Click()
'    Dim strInfo$
'    strInfo = "QLunch | 快速启动 V" & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf & vbCrLf & _
'        "  作者:sysdzw" & vbCrLf & _
'        "  主页:https://blog.csdn.net/sysdzw" & vbCrLf & _
'        "  Q  Q:171977759" & vbCrLf & _
'        "  邮箱:sysdzw@163.com" & vbCrLf & vbCrLf & _
'        "2023-05-18"
'        MsgBox strInfo, vbInformation
        
    frmAbout.Show 1
End Sub

Private Sub Command5_Click()
    Shell "explorer /select,""" & List2.List(List1.ListIndex) & """", 1
End Sub

Private Sub Form_Click()
    MsgBox getListAllItem(List2), , App.Title
End Sub

Private Sub Form_Load()
    Dim w As New clsWindow
    If w.GetWindowByTitleEx("快速启动 v", 0, , , , DisplayedWindow).hWnd <> 0 Then
        w.Show
        w.Focus
        End
    End If

    Me.Caption = proName
    Dim i%, v, s$
    strSetFile = strAppPath & "设置.txt"
    s = fileStr(strSetFile)
    If s <> "" Then
        v = Split(s, vbCrLf)
        For i = 0 To UBound(v)
            List1.AddItem getFileName(v(i))
            List2.AddItem v(i)
        Next
    End If
End Sub

Private Sub Form_Resize()
On Error GoTo Err1
    List1.Move 45, List1.Top, Me.ScaleWidth - 90, Me.ScaleHeight - lblTip.Height - List1.Top - 90
    lblTip.Move 45, Me.ScaleHeight - lblTip.Height - 45, Me.ScaleWidth - 90
    Command4.Left = Me.ScaleWidth - Command4.Width - 90
    Command4.Visible = Command4.Left - (lblMsg.Left + lblMsg.Width) > 45
Err1:
End Sub

Private Sub List1_Click()
    lblTip.Caption = List2.List(List1.ListIndex)
End Sub

Private Sub List1_DblClick()
    openFileDbl Me.hWnd, List2.List(List1.ListIndex)
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call List1_DblClick
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    intMoveStart = List1.ListIndex
    List1.MousePointer = 5
    
    Label1.Caption = List1.List(List1.ListIndex)
    Text1.Text = " " & List1.List(List1.ListIndex)
    Text1.Width = Label1.Width + 200 '利用下Label1设置Text1的宽度
    
    Text1.Move X, Y - 150
    Text1.Visible = True
    isMove = True
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If isMove Then Text1.Move X, Y

    Static intSelectedIndex As Integer
    If List1.ListIndex <> intSelectedIndex Then
        intSelectedIndex = List1.ListIndex
'        Me.Caption = List1.ListIndex & " " & Now
    End If
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    isMove = False
    Text1.Visible = False
    
    List1.MousePointer = 0
    intMoveEnd = List1.ListIndex
    If intMoveStart = intMoveEnd Then Exit Sub
    List1.AddItem List1.List(intMoveStart), intMoveEnd
    List2.AddItem List2.List(intMoveStart), intMoveEnd
    List1.RemoveItem IIf(intMoveEnd > intMoveStart, intMoveStart, intMoveStart + 1)
    List2.RemoveItem IIf(intMoveEnd > intMoveStart, intMoveStart, intMoveStart + 1)
    List1.Selected(intMoveEnd) = True
End Sub

Private Sub List1_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDragFile As String
    On Error GoTo Err1
    If Data.GetFormat(1) Then Exit Sub
    
    strDragFile = Data.Files.Item(Data.Files.Count)
    If InStr(getListAllItem(List2), strDragFile) = 0 Then
        List1.AddItem getFileName(strDragFile)
        List2.AddItem strDragFile
        writeToFile strSetFile, getListAllItem(List2)
    Else
        MsgBox "该项目已存在，请勿重复添加", vbInformation
    End If
    
Err1:
    Exit Sub
End Sub

Function getListAllItem(List1 As ListBox) As String
    Dim i%, s$
    For i = 0 To List1.ListCount - 1
        s = s & List1.List(i) & vbCrLf
    
    Next
    If Right(s, 2) = vbCrLf Then s = Left(s, Len(s) - 2)
    getListAllItem = s
End Function
Function getFileName(ByVal str1$) As String
    Dim v
    v = Split(str1, "\")
    getFileName = v(UBound(v))
End Function

