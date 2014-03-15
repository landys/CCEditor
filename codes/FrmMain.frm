VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "CCEditor"
   ClientHeight    =   6225
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10545
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   2400
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   5925
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13000
            MinWidth        =   882
            Key             =   "TipText"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1429
            MinWidth        =   882
            TextSave        =   "2005-8-13"
            Key             =   "Date"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "15:24"
            Key             =   "Time"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "INS"
            Key             =   "Insert"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            TextSave        =   "CAPS"
            Key             =   "Caps"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
            Key             =   "Nums"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PicNavigation 
      Align           =   3  'Align Left
      AutoSize        =   -1  'True
      Height          =   5520
      Left            =   0
      ScaleHeight     =   5460
      ScaleWidth      =   1995
      TabIndex        =   1
      Top             =   405
      Width           =   2055
      Begin VB.FileListBox filEditor 
         Height          =   3210
         Left            =   0
         TabIndex        =   5
         Top             =   2160
         Width           =   1935
      End
      Begin VB.DirListBox dirEditor 
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.DriveListBox drvEditor 
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList imglstToolbar 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0CCA
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0DD4
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0EDE
            Key             =   "save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":0FE8
            Key             =   "undo"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":10F2
            Key             =   "redo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11FC
            Key             =   "cut"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1306
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1410
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":151A
            Key             =   "delete"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1624
            Key             =   "help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   714
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      ImageList       =   "imglstToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "new"
            Object.ToolTipText     =   "New File"
            ImageKey        =   "new"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "open"
            Object.ToolTipText     =   "Open File"
            ImageKey        =   "open"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "save"
            Object.ToolTipText     =   "Save File"
            ImageKey        =   "save"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "undo"
            Object.ToolTipText     =   "Undo"
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "redo"
            Object.ToolTipText     =   "Redo"
            ImageKey        =   "redo"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cut"
            Object.ToolTipText     =   "Cut"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "copy"
            Object.ToolTipText     =   "Copy"
            Object.Tag             =   "Copy"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "paste"
            Object.ToolTipText     =   "Paste"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Object.ToolTipText     =   "Help"
            ImageKey        =   "help"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRecent 
         Caption         =   "Recent Files"
         Begin VB.Menu mnuFileRec 
            Caption         =   "Empty"
            Index           =   0
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File1"
            Index           =   1
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File2"
            Index           =   2
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File3"
            Index           =   3
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File4"
            Index           =   4
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File5"
            Index           =   5
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File6"
            Index           =   6
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File7"
            Index           =   7
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File8"
            Index           =   8
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File9"
            Index           =   9
            Visible         =   0   'False
         End
         Begin VB.Menu mnuFileRec 
            Caption         =   "Recent File10"
            Index           =   10
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuFileSeperator3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSTatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewNavigation 
         Caption         =   "&Navigation"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsClearRec 
         Caption         =   "Clear Recent Files"
      End
      Begin VB.Menu mnuOptionsNumber 
         Caption         =   "Set Recent Files &Maximum..."
      End
      Begin VB.Menu mnuOptionsSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsTools 
         Caption         =   "&Tools"
         Begin VB.Menu mnuOptionsToolsCalculator 
            Caption         =   "&Calculator"
         End
         Begin VB.Menu mnuOptionsToolsClipboard 
            Caption         =   "&Clipboard"
         End
         Begin VB.Menu mnuOptionsToolsIE 
            Caption         =   "&Internet Explorer"
         End
         Begin VB.Menu mnuOptionsToolsPaint 
            Caption         =   "&Paint"
         End
      End
      Begin VB.Menu mnuSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsGames 
         Caption         =   "&Games"
         Begin VB.Menu mnuOptionsGamesHearts 
            Caption         =   "&Hearts"
         End
         Begin VB.Menu mnuOptionsGamesMiner 
            Caption         =   "&Miner"
         End
         Begin VB.Menu mnuOptionsGamesSol 
            Caption         =   "&Sol"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "&Help Topic..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAboutCCEditor 
         Caption         =   "&About CCEditor..."
      End
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************'
'*     本父窗体模块实现该程序的界面布局等功能。主要包括目录文件导航实现、工具栏和状 *'
'* 态栏的实现等。                                                               *'
'*******************************************************************************'

Option Explicit

' 声明连接帮助信息的 Windows 函数
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, _
    ByVal dwData As Long) As Long

' DirlistBox 控件中目录改变

Private Sub dirEditor_Change()
    On Error GoTo Dealerror ' 可能文件夹不存在错误
    
    Dim dPath As String  ' 保存路径
    
    dPath = filEditor.Path  ' 保存路径，以便恢复
    filEditor.Path = dirEditor.Path ' 设置 FilelistBox 控件路径
    Exit Sub
Dealerror:  ' 提示，并恢复原路径
    MsgBox ("Directory Not Available")
    filEditor.Path = dPath
    dirEditor.Path = dPath
End Sub

Private Sub dirEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Direction change on navigation"    ' 当前状态提示
End Sub

' DrivelistBox 控件中目录改变

Private Sub drvEditor_Change()
    On Error GoTo Dealerror ' 可能分区不可用错误，比如未插入光盘等
    
    FrmMain.StatusBar1.Panels(1).Text = "Drive change on navigation."    ' 当前状态提示
    
    Dim dPath As String  ' 保存路径
    
    dPath = dirEditor.Path  ' 保存路径，以便恢复
    dirEditor.Path = drvEditor.Drive ' 设置 DirlistBox 控件分区符
    Exit Sub
Dealerror:  ' 提示，并恢复原分区
    MsgBox ("Drive Not Available")
    dirEditor.Path = dPath
    drvEditor.Drive = dPath
End Sub

' 双击打开文件

Private Sub filEditor_DblClick()
    On Error Resume Next
    
    Dim strFileName As String
    
    strFileName = filEditor.Path + "\" + filEditor.FileName
    NewFile  ' 打开新文档窗体
    Me.ActiveForm.rtfText.LoadFile strFileName  ' 加载文件
    Me.ActiveForm.Caption = strFileName
    Me.ActiveForm.m_Modified = False
    
    UpdateFileMenu (strFileName)    ' 更新最近显示文件列表，该函数在模块 mdlFile 中
End Sub


Private Sub filEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Files change on navigation."    ' 当前状态提示
End Sub

' 窗体加载初始化

Private Sub MDIForm_Load()
    ' 初始化局变量
    gFindString = ""
    gReplaceString = ""
    gFindCase = False
    gFindDirection = 1
    gCurPos = 0
    gCountFrmChild = 0   ' 附值全局变量，子窗体个数
    
    ' 若注册表中无此程序信息，则写入
    If GetSetting(App.Title, "Recent Files", "Maximum Of Recent Files") = Empty Then
        SaveSetting App.Title, "Recent Files", "Maximum Of Recent Files", 6
    End If
    ' 从注册表获取最近打开文档最大值
    gMaxRecentFiles = GetSetting(App.Title, "Recent Files", "Maximum Of Recent Files")
    NewFile ' 新建文件，该函数在模块 mdlFile 中
    
    GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
    FrmMain.StatusBar1.Panels(1).Text = "For Help, Press F1"    ' 当前状态提示
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "For Help, Press F1"    ' 当前状态提示
End Sub

Private Sub MDIForm_Resize()
    PicNavigation_Resize    ' 包容目录文件导航的图片框调整大小
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload frmFind  ' 关闭查找无模式对话框
End Sub

Private Sub mnuFileExit_Click()
    Unload Me   ' 退出
End Sub

' 创建新文件

Private Sub mnuFileNew_Click()
    NewFile ' 该函数在模块 mdlFile 中
End Sub

' 打开文件

Private Sub mnuFileOpen_Click()
    OpenFile ' 该函数在模块 mdlFile 中
End Sub

' 最近打开文件菜单

Private Sub mnuFileRec_Click(index As Integer)
    On Error GoTo Dealerror
    
    Dim strFileName As String   ' 文件名
    Dim strKey As String    ' 在注册表中的关键字项
    
    If index = 0 Then ' 点击 Empty 菜单直接结束过程
        Exit Sub
    End If
    
    strKey = "Recent File" & index  ' 从Windows 注册表中的应用程序项目返回注册表项设置值
    strFileName = GetSetting(App.Title, "Recent Files", strKey)
    NewFile  ' 打开新文档窗体，该函数在模块 mdlFile 中
    Me.ActiveForm.rtfText.LoadFile strFileName
    Me.ActiveForm.Caption = strFileName
    UpdateFileMenu (strFileName)
    
    Me.ActiveForm.m_Modified = False ' 设置标志为未修改
    Exit Sub
    
Dealerror:  ' 文件不存在错误处理，提示信息，并删除相应注册表项
    Unload Me.ActiveForm   ' 由于打开文件前新建了窗体，在此先删除
    MsgBox "The File Not Exist"
    DeleteRecentFile index  ' 删除相应注册表项，该函数在模块 mdlFile 中
    GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
End Sub

Private Sub mnuHelpAboutCCEditor_Click()
    frmAbout.Show 1 ' 显示 About 对话框
End Sub

Private Sub mnuHelpTopic_Click()
    WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' 显示帮助信息
End Sub

' 清除最近保存文件列表菜单

Private Sub mnuOptionsClearRec_Click()
    DeleteAllRecentFiles    ' 清除最近保存文件列表菜单，该函数在模块 mdlFile 中
    GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
End Sub

' 最近打开文档列表最大数目设置，不能超过9

Private Sub mnuOptionsNumber_Click()
    On Error Resume Next
    
    Dim intN As Integer ' 输入的设置大小
    Dim strN As String  ' 从输入对话框输入的文本
    Dim intCount As Integer ' 输入错误计数
    Dim i As Integer    ' For 循环索引计数
    
    intCount = 0
    ' 输入对话框
    strN = InputBox("Input maximum of recent files showed on menu. The number should be 0--9.")
    intN = Asc(strN) - Asc("0") ' 得到数字大小
    If Len(strN) > 1 Then   ' 若输入多个字符，肯定为错
        intN = -1
    End If
    
    While intN < 0 Or intN > 10 ' 出错
        intCount = intCount + 1 ' 输入错误计数加 1
        If (intCount = 3) Then  ' 连续三次输入错误，消息提示
            MsgBox "Input wrong for three times! Maximum of recent files not changed"
            Exit Sub
        End If
        ' 提示重输入
        MsgBox "Input wrong! Please input again."
        intN = InputBox("Input maximum of recent files showed on menu. The number should be 0--9.")
    Wend
    If gMaxRecentFiles = intN Then
        Exit Sub
    ElseIf gMaxRecentFiles > intN Then  ' 若要设置的值比原来的小，删除多余注册表项，并更新最近打开文件列表
        For i = gMaxRecentFiles To intN + 1 Step -1
            DeleteRecentFile i  ' 删除相应注册表项，该函数在模块 mdlFile 中
        Next i
        GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
    End If
    
    gMaxRecentFiles = intN  ' 设置最近打开文件的最大数值
    ' 保存注册表项
    SaveSetting App.Title, "Recent Files", "Maximum Of Recent Files", gMaxRecentFiles
    Exit Sub
End Sub

' 显示或隐藏程序左边的目录文件导航

Private Sub mnuViewNavigation_Click()
    mnuViewNavigation.Checked = Not mnuViewNavigation.Checked
    Me.PicNavigation.Visible = Not Me.PicNavigation.Visible
End Sub

' 显示或隐藏状态栏

Private Sub mnuViewStatusbar_Click()
    mnuViewSTatusbar.Checked = Not mnuViewSTatusbar.Checked
    Me.StatusBar1.Visible = Not Me.StatusBar1.Visible
End Sub

' 显示或隐藏工具栏

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Me.Toolbar1.Visible = Not Me.Toolbar1.Visible
End Sub

Private Sub PicNavigation_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Look on navigation."    ' 当前状态提示
End Sub

' 包容目录文件导航的图片框调整大小
    
Private Sub PicNavigation_Resize()
    On Error Resume Next
    
    drvEditor.Move 0, 0
    dirEditor.Move 0, (drvEditor.Height + 20), drvEditor.Width, (PicNavigation.Height - (drvEditor.Height + 20)) / 3
    filEditor.Move 0, drvEditor.Height + dirEditor.Height + 40, drvEditor.Width, PicNavigation.Height - dirEditor.Height - drvEditor.Height - 30
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Cursor On Statusbar"    ' 当前状态提示
End Sub

' 实现工具栏

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    FrmMain.StatusBar1.Panels(1).Text = Button.ToolTipText  ' 当前状态提示
    Select Case Button.Key
        Case "new"
            NewFile ' 创建新文件，该函数在模块 mdlFile 中
        Case "open"
            OpenFile    ' 打开文件，该函数在模块 mdlFile 中
        Case "save"
            SaveFile    ' 保存文件，该函数在模块 mdlFile 中
        Case "undo"
            Undo    ' 撤销
        Case "redo"
            Redo    ' 重做
        Case "cut"
            CutText    ' 剪切，该函数在模块 mdlEdit 中
        Case "copy"
            CopyText    ' 拷贝，该函数在模块 mdlEdit 中
        Case "paste"
            PasteText   ' 粘贴，该函数在模块 mdlEdit 中
        Case "delete"
            DeleteText  ' 删除，该函数在模块 mdlEdit 中
        Case "help"
            WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' 显示帮助信息
    End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Cursor On Toolbar"    ' 当前状态提示
End Sub

' 附属剪贴板工具

Private Sub mnuOptionsToolsClipboard_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\clipbrd.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 附属计算器工具

Private Sub mnuOptionsToolsCalculator_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\calc.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 附属画图板工具

Private Sub mnuOptionsToolsPaint_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mspaint.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 附属游戏扫雷

Private Sub mnuOptionsGamesMiner_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\winmine.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 附属游戏红心接龙

Private Sub mnuOptionsGamesHearts_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mshearts.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"

End Sub

' 附属游戏纸牌

Private Sub mnuOptionsGamesSol_Click()

    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\sol.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 附属 IE 浏览器工具

Private Sub mnuOptionsToolsIE_Click()
    On Error GoTo Dealerror
    
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub
