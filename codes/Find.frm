VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Find"
   ClientHeight    =   2190
   ClientLeft      =   2655
   ClientTop       =   3585
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   146
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   337
   Begin VB.CommandButton cmdReplaceAll 
      Caption         =   "Replace &All"
      Height          =   375
      Left            =   3720
      TabIndex        =   11
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdReplace 
      Caption         =   "&Replace"
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox txtReplace 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   840
      Width           =   2175
   End
   Begin VB.Frame FraDirection 
      Caption         =   "Direction"
      Height          =   735
      Left            =   1560
      TabIndex        =   3
      Top             =   1320
      Width           =   2052
      Begin VB.OptionButton optDirection 
         Caption         =   "&Down"
         Height          =   252
         Index           =   1
         Left            =   960
         TabIndex        =   5
         ToolTipText     =   "Search to End of Document"
         Top             =   240
         Value           =   -1  'True
         Width           =   852
      End
      Begin VB.OptionButton optDirection 
         Caption         =   "&Up"
         Height          =   252
         Index           =   0
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Search to Beginning of Document"
         Top             =   240
         Width           =   612
      End
   End
   Begin VB.CheckBox chkCase 
      Caption         =   "Match &Case"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Case Sensitivity"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txtFind 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      ToolTipText     =   "Text to Find"
      Top             =   240
      Width           =   2175
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   372
      Left            =   3720
      TabIndex        =   7
      ToolTipText     =   "Return to Notepad"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find"
      Default         =   -1  'True
      Height          =   372
      Left            =   3720
      TabIndex        =   6
      ToolTipText     =   "Start Search"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label lblReplace 
      Caption         =   "R&eplace with:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblFind 
      Caption         =   "Fi&nd What:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************'
'*     查找对话框模块，对话框中的控件响应，实现查找替换功能。主要功能为查找、查找下一个，替换 *'
'* 及一次性全部替换，另外还包括是否区分大小写、搜索方向设定等。                             *'
'****************************************************************************************'

Option Explicit

' 声明设定窗口位置 Windows 函数，可用来使该窗口始终位于最上层
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


' CheckBox，设定是否区分大小

Private Sub chkCase_Click()
    gFindCase = chkCase.Value   ' 对全局变量附值
End Sub

' Cancel 退出

Private Sub cmdCancel_Click()
    ' 对全局变量附值，供下次使用
    gFindString = txtFind.Text
    gReplaceString = txtReplace.Text
    gFindCase = chkCase.Value
    
    Unload frmFind  ' 退出
End Sub

' 查找按钮

Private Sub cmdFind_Click()
    gFindString = txtFind.Text   ' 对全局变量附值
    cmdFind.Caption = "Find Next"
    FindIt  ' 此函数在模块 mdlEdit 中，实现查找
    
    gFirstTime = False  ' 标志已查找过
End Sub

' 替换按钮

Private Sub cmdReplace_Click()
    gFindString = txtFind.Text   ' 对全局变量附值
    gReplaceString = txtReplace.Text
    ReplaceIt  ' 此函数在模块 mdlEdit 中，实现替换
    
    gFirstTime = False  ' 标志已查找过
End Sub

' 替换全部按钮

Private Sub cmdReplaceAll_Click()
    gFindString = txtFind.Text   ' 对全局变量附值
    gReplaceString = txtReplace.Text
    
    ReplaceAll  ' 此函数在模块 mdlEdit 中定义，替换所有要查找的字符串
    
    gFirstTime = False  ' 标志已查找过
End Sub

' 窗体加载时初始化

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 510, 195, 345, 180, 0 ' 该窗口总是显示上其他窗口之上
    
    cmdFind.Enabled = False ' Disable 查找按钮
    cmdReplace.Enabled = False  ' Disable 替换按钮
    cmdReplaceAll.Enabled = False   ' Disable 替换全部按钮

    optDirection(gFindDirection).Value = 1  ' 读取搜索方向，设置
End Sub

' OptionButton 设置搜索方向

Private Sub optDirection_Click(index As Integer)
    gFindDirection = index  ' 设置搜索方向
End Sub

' 输入要查找的文本内容改变

Private Sub txtFind_Change()
    gFirstTime = True   ' 对全局变量附值
    
    cmdFind.Caption = "Find"    ' 设置查找按钮显示信息为 Find
    
    If txtFind.Text = "" Then   ' 查找 TextBox 为空时，Disable 各相关按钮
        cmdFind.Enabled = False
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else   ' 查找 TextBox 不为空时，Enable 各相关按钮
        cmdFind.Enabled = True
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
End Sub

