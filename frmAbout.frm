VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "关于"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5670
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5670
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command1 
      Caption         =   "确 定"
      Height          =   450
      Left            =   1680
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "github:"
      Height          =   180
      Left            =   960
      TabIndex        =   11
      Top             =   1080
      Width           =   630
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "https://github.com/sysdzw/QLunch"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1680
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   1080
      Width           =   2880
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "2023-05-18"
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   2040
      Width           =   900
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "171977759@qq.com"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   1560
      Width           =   1440
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "邮 箱:"
      Height          =   180
      Left            =   1050
      TabIndex        =   7
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "171977759"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1680
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1320
      Width           =   810
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Q  Q:"
      Height          =   180
      Left            =   1140
      TabIndex        =   5
      Top             =   1320
      Width           =   450
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "https://sysdzw.blog.csdn.net/?type=blog"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   180
      Left            =   1680
      MouseIcon       =   "frmAbout.frx":015E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Top             =   840
      Width           =   3510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "博 客:"
      Height          =   180
      Left            =   1050
      TabIndex        =   3
      Top             =   840
      Width           =   540
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "作 者:sysdzw"
      Height          =   180
      Left            =   1035
      TabIndex        =   1
      Top             =   600
      Width           =   1080
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   555
      Left            =   240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "软件名称"
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   720
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmMain.Icon
    Image1 = frmMain.Icon
    Label6.MouseIcon = Label4.MouseIcon
    Label8.MouseIcon = Label4.MouseIcon
    
    Label1.Caption = App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision
    Label4.Caption = "https://sysdzw.blog.csdn.net/?type=blog"
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub Label10_Click()
    ShellExecute hwnd, "open", Label10.Caption, "", "", 1
End Sub

Private Sub Label4_Click()
    ShellExecute hwnd, "open", Label4.Caption, "", "", 1
End Sub

Private Sub Label6_Click()
    ShellExecute 0, "", "tencent://message?uin=" & Label6.Caption & "&Site=好人&Menu=yes", "", "", 5
End Sub

Private Sub Label8_Click()
    ShellExecute hwnd, "open", "mailto:" & Label8.Caption, "", "", 1
End Sub

