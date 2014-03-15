VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frmChild 
   BackColor       =   &H80000005&
   Caption         =   "Noname"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8655
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   12
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmChild.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   4335
   ScaleWidth      =   8655
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7435
      _Version        =   327681
      Enabled         =   -1  'True
      ScrollBars      =   3
      OLEDropMode     =   1
      TextRTF         =   $"FrmChild.frx":0CCA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Fixedsys"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
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
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSeperator2 
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
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "&Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "&Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select &ALL"
         Shortcut        =   ^A
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewStatusbar 
         Caption         =   "&Statusbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewNavigation 
         Caption         =   "Navi&gation"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchFind 
         Caption         =   "&Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSearchFindNext 
         Caption         =   "Find &Next"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuSearchReplace 
         Caption         =   "&Replace..."
         Shortcut        =   ^{F3}
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptionsLeft 
         Caption         =   "Align To &Left"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuOptionsRight 
         Caption         =   "Align To &Right"
      End
      Begin VB.Menu mnuOptionsCenter 
         Caption         =   "Align To &Center"
      End
      Begin VB.Menu mnuOptionsSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsFont 
         Caption         =   "&Font..."
      End
      Begin VB.Menu mnuOptionsColor 
         Caption         =   "&Color"
         Begin VB.Menu mnuOptionsColorText 
            Caption         =   "&Text Color..."
         End
         Begin VB.Menu mnuOptionsColorChar 
            Caption         =   "&Char Color..."
         End
         Begin VB.Menu mnuOptionsColorNum 
            Caption         =   "&Num Color..."
         End
      End
      Begin VB.Menu mnuOptionsSeperator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOptionsClearRec 
         Caption         =   "&Clear Recent Files"
      End
      Begin VB.Menu mnuOptionsNumber 
         Caption         =   "Set Recent Files &Maximum..."
      End
      Begin VB.Menu mnuOptionsSeperator3 
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
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHor 
         Caption         =   "Tile &Horizontally"
      End
      Begin VB.Menu mnuWindowTileVer 
         Caption         =   "Tile &Vertically"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
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
Attribute VB_Name = "FrmChild"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************'
'*     本子窗体模块实现该程序的主要功能。最主要的为控件 RichTextBox 的实现，包括输 *'
'* 入、读文件、写文件及菜单上的所有响应。另外比较主要的功能实现有：工具栏和状态栏的 *'
'* 实现，最近打开文件菜单列表的实现，对新建或打开的文件进行编辑（插入、修改、删除） *'
'* 的实现，剪切板、查找、替换功能的实现，多步撤销功能的实现，文件中的数字、字符应能 *'
'* 用不同的颜色显示及设置的实现，在线帮助功能的实现，提供辅助工具功能的实现等。     *'
'*******************************************************************************'

Option Explicit

Public m_Modified As Boolean    ' 标志文件是否已修改
Private m_CharColor As Long  ' 字符颜色
Private m_NumColor As Long   ' 数字字符颜色
Private m_TextChange As Boolean  ' 在 RichTextBox 文字改变时使用
                                 ' 主要用于字符数字以不同颜色显示的功能实现

' 声明连接帮助信息的 Windows 函数
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, _
    ByVal dwData As Long) As Long


' 窗口被激活

Private Sub Form_Activate()
    GetRecentFiles  ' 该函数在模块 mdlFile 中，更新菜单中最近打开文档列表
End Sub

' 窗体加载初始化

Private Sub Form_Load()
    gCountFrmChild = gCountFrmChild + 1   ' 附值全局变量，子窗体个数加 1
    Form_Resize ' 实现窗体内的各控件布局
    m_CharColor = 0  ' 字符颜色
    m_NumColor = 0   ' 数字字符颜色
End Sub

' 窗体加载前初始化

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intResult As Integer
    If m_Modified Then
    '倘若文件被修改过则提示是否保存
        intResult = MsgBox("Save modified file to " & Me.Caption, _
            vbYesNoCancel + vbQuestion, "CCEditor")
        If intResult = vbCancel Then
            Cancel = 1
        ElseIf intResult = vbYes Then
            mnuFileSave_Click   '保存文件后退出
            Cancel = 0
        Else
            Cancel = 0  '不保存退出
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    ' 子窗口大小变化时，文本框跟着变化
    rtfText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    rtfText.RightMargin = rtfText.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gCountFrmChild = gCountFrmChild - 1 ' 附值全局变量，子窗体个数减 1
End Sub

Private Sub mnuEditCopy_Click()
    CopyText    ' 拷贝，该函数在模块 mdlEdit 中
End Sub

Private Sub mnuEditCut_Click()
    CutText    ' 剪切，该函数在模块 mdlEdit 中
End Sub

Private Sub mnuEditDelete_Click()
    DeleteText  ' 删除，该函数在模块 mdlEdit 中
End Sub

Private Sub mnuEditPaste_Click()
    PasteText   ' 粘贴，该函数在模块 mdlEdit 中
End Sub

Private Sub mnuEditRedo_Click()
    Redo    ' 重做
End Sub

' 选中所有内容

Private Sub mnuEditSelectAll_Click()
    rtfText.SelStart = 0
    rtfText.SelLength = Len(rtfText.TextRTF)
End Sub

Private Sub mnuEditUndo_Click()
    Undo    ' 撤销
End Sub

Private Sub mnuFileClose_Click()
    Unload Me   ' 关闭文件
End Sub

Private Sub mnuFileExit_Click()
    Unload FrmMain  ' 关闭程序
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
    rtfText.LoadFile strFileName
    Me.Caption = strFileName
    UpdateFileMenu (strFileName)
    Me.SetFocus
    
    m_Modified = False ' 设置标志为未修改
    Exit Sub
    
Dealerror:  ' 文件不存在错误处理，提示信息，并删除相应注册表项
    Unload Me   ' 由于打开文件前新建了窗体，在此先删除
    MsgBox "The File Not Exist"
    DeleteRecentFile index  ' 删除相应注册表项，该函数在模块 mdlFile 中
    GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveFileAs  ' 另存为，该函数在模块 mdlFile 中
End Sub

Private Sub mnuHelpAboutCCEditor_Click()
    frmAbout.Show 1 ' 显示 About 对话框
End Sub

Private Sub mnuHelpTopic_Click()
    WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' 显示帮助信息
End Sub

' 设置为中间对齐

Private Sub mnuOptionsCenter_Click()
    
    mnuOptionsLeft.Checked = False
    mnuOptionsRight.Checked = False
    mnuOptionsCenter.Checked = True
    rtfText.SelAlignment = rtfCenter
End Sub

' 清除最近保存文件列表菜单

Private Sub mnuOptionsClearRec_Click()
    DeleteAllRecentFiles    ' 清除最近保存文件列表菜单，该函数在模块 mdlFile 中
    GetRecentFiles  ' 更新最近打开文件列表菜单，该函数在模块 mdlFile 中
End Sub

' 字符颜色设置

Private Sub mnuOptionsColorChar_Click()
    On Error Resume Next
    
    Dim lgSelStart As Long
    Dim lgSelLength As Long
    Dim i As Long
    
    m_TextChange = False    ' 标志，这样在 rtfText_Change 过程被调用时避免运行无关代码
    
    ' 颜色通用对话框初始化
    FrmMain.dlgCommonDialog.Color = m_CharColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    ' 这里通过一个个字符判断实现
    If m_CharColor <> FrmMain.dlgCommonDialog.Color Then    ' 颜色已改变
        m_CharColor = FrmMain.dlgCommonDialog.Color
        If rtfText.SelText <> "" Then   ' 有文本选中
            lgSelStart = rtfText.SelStart   ' 保存选中文本的位置，以便恢复
            lgSelLength = rtfText.SelLength   ' 保存选中文本的长度，以便恢复
            For i = 0 To lgSelLength - 1    ' 逐字符比较，改变颜色
                rtfText.SelStart = lgSelStart + i
                rtfText.SelLength = 1
                If Not (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then
                    ' 非数字字符
                    rtfText.SelColor = m_CharColor
                End If
            Next
            rtfText.SelStart = lgSelStart ' 恢复选中文本的位置
            rtfText.SelLength = lgSelLength   ' 恢复选中文本的长度
        End If
    End If
    
    m_TextChange = True ' 恢复标志
End Sub

' 数字颜色设置

Private Sub mnuOptionsColorNum_Click()
    On Error Resume Next
    
    Dim lgSelStart As Long
    Dim lgSelLength As Long
    Dim i As Long
    
    m_TextChange = False    ' 标志，这样在 rtfText_Change 过程被调用时避免运行无关代码
    
    ' 颜色通用对话框初始化
    FrmMain.dlgCommonDialog.Color = m_NumColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    ' 这里通过一个个字符判断实现
    If m_NumColor <> FrmMain.dlgCommonDialog.Color Then    ' 颜色已改变
        m_NumColor = FrmMain.dlgCommonDialog.Color
        If rtfText.SelText <> "" Then   ' 有文本选中
            lgSelStart = rtfText.SelStart   ' 保存选中文本的位置，以便恢复
            lgSelLength = rtfText.SelLength   ' 保存选中文本的长度，以便恢复
            For i = 0 To lgSelLength - 1    ' 逐字符比较，改变颜色
                rtfText.SelStart = lgSelStart + i
                rtfText.SelLength = 1
                If (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then ' 非数字字符
                    rtfText.SelColor = m_NumColor
                End If
            Next
            rtfText.SelStart = lgSelStart ' 恢复选中文本的位置
            rtfText.SelLength = lgSelLength   ' 恢复选中文本的长度
        End If
    End If
    
    m_TextChange = True ' 恢复标志
End Sub

' 文本颜色设置

Private Sub mnuOptionsColorText_Click()
    m_TextChange = False    ' 标志，这样在 rtfText_Change 过程被调用时避免运行无关代码
    
    ' 颜色通用对话框初始化
    FrmMain.dlgCommonDialog.Color = m_CharColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    m_CharColor = FrmMain.dlgCommonDialog.Color
    m_NumColor = m_CharColor
    rtfText.SelColor = m_CharColor  ' 文本颜色设置
    
    m_TextChange = True ' 恢复标志
End Sub

' 设置字体

Private Sub mnuOptionsFont_Click()
    On Error Resume Next
    
    m_TextChange = False    ' 标志，这样在 rtfText_Change 过程被调用时避免运行无关代码
    
    ' 用通用对话框设置字体
    With FrmMain.dlgCommonDialog
        .FontName = rtfText.Font.Name
        .FontBold = rtfText.Font.Bold
        .FontItalic = rtfText.Font.Italic
        .FontSize = rtfText.Font.Size
        .FontStrikethru = rtfText.Font.Strikethrough
        .FontUnderline = rtfText.Font.Underline
        .Flags = cdlCFBoth + cdlCFEffects
        .ShowFont
        .CancelError = True
    End With
    With rtfText
        .SelFontName = FrmMain.dlgCommonDialog.FontName
        .SelFontSize = FrmMain.dlgCommonDialog.FontSize
        .SelBold = FrmMain.dlgCommonDialog.FontBold
        .SelItalic = FrmMain.dlgCommonDialog.FontItalic
        .SelStrikeThru = FrmMain.dlgCommonDialog.FontStrikethru
        .SelUnderline = FrmMain.dlgCommonDialog.FontUnderline
    End With
    
    m_TextChange = True ' 恢复标志
End Sub

' 附属游戏红心接龙

Private Sub mnuOptionsGamesHearts_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mshearts.exe", 1
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

' 附属游戏纸牌

Private Sub mnuOptionsGamesSol_Click()

    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\sol.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' 设置为文本左对齐
    
Private Sub mnuOptionsLeft_Click()
    mnuOptionsLeft.Checked = True
    mnuOptionsRight.Checked = False
    mnuOptionsCenter.Checked = False
    rtfText.SelAlignment = rtfLeft
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

' 设置为文本右对齐

Private Sub mnuOptionsRight_Click()
    mnuOptionsLeft.Checked = False
    mnuOptionsRight.Checked = True
    mnuOptionsCenter.Checked = False
    rtfText.SelAlignment = rtfRight
End Sub

' 附属剪贴板工具

Private Sub mnuOptionsToolsClipboard_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\clipbrd.exe", 1
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

' 查找

Private Sub mnuSearchFind_Click()
    ' 查找的初始值为上次查找的值或者文件选中的字符
    If Me.rtfText.SelText <> "" Then
        frmFind.txtFind.Text = Me.rtfText.SelText
    Else
        frmFind.txtFind.Text = gFindString
    End If
    
    gFirstTime = True  ' 全局变量，表明首次查找
    
    If (gFindCase) Then ' 设置是否区分大小写
        frmFind.chkCase = 1
    End If
    
    frmFind.Show 0  ' 显示查找对话框
End Sub

' 查找下一个
Private Sub mnuSearchFindNext_Click()
    If Len(gFindString) > 0 Then    ' 若查找的文本不为空
        gFirstTime = False
        FindIt  ' 实现查找，该函数在模块 mdlEdit 中定义
    Else    ' 若查找为空
        mnuSearchFind_Click ' 若查找为空
    End If
End Sub

' 替换

Private Sub mnuSearchReplace_Click()
    frmFind.txtReplace.Text = gReplaceString
    mnuSearchFind_Click
End Sub

' 显示或隐藏程序左边的目录文件导航

Private Sub mnuViewNavigation_Click()
    mnuViewNavigation.Checked = Not mnuViewNavigation.Checked
    FrmMain.PicNavigation.Visible = Not FrmMain.PicNavigation.Visible
End Sub

' 显示或隐藏状态栏

Private Sub mnuViewStatusbar_Click()
    mnuViewSTatusbar.Checked = Not mnuViewSTatusbar.Checked
    FrmMain.StatusBar1.Visible = Not FrmMain.StatusBar1.Visible
End Sub

' 显示或隐藏工具栏

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    FrmMain.Toolbar1.Visible = Not FrmMain.Toolbar1.Visible
End Sub

' 将所有MDI最小化的子窗体图标进行重排
    
Private Sub mnuWindowArrangeIcons_Click()
    FrmMain.Arrange vbArrangeIcons
End Sub

' 将所有MDI非最小化的子窗体层叠排放

Private Sub mnuWindowCascade_Click()
    FrmMain.Arrange vbCascade
End Sub

' 将所有MDI非最小化的子窗体水平排放

Private Sub mnuWindowTileHor_Click()
    FrmMain.Arrange vbTileHorizontal
End Sub

' 将所有MDI非最小化的子窗体垂直排放

Private Sub mnuWindowTileVer_Click()
    FrmMain.Arrange vbTileVertical
End Sub

' 保存文件

Private Sub mnuFileSave_Click()
    SaveFile     ' 该函数在模块 mdlFile 中
End Sub

Private Sub rtfText_Change()
    ' 如果文本有变化，则设置修改标志
    m_Modified = True

    ' 字符输入后判断其是数字还是非数字，再以不同颜色显示
    If m_TextChange And m_CharColor <> 0 And m_NumColor <> 0 Then
        rtfText.SelStart = rtfText.SelStart - 1 ' 选中这个字符，再改为相应颜色
        rtfText.SelLength = 1
        If (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then
            rtfText.SelColor = m_NumColor
        Else
            rtfText.SelColor = m_CharColor
        End If
        
        rtfText.SelStart = rtfText.SelStart + 1 '恢复光标
        rtfText.SelLength = 0   ' 取消选中
    End If
    
    If Not trapUndo Then  ' 操作跟踪被禁止
        Exit Sub
    End If

    Dim newElement As New UndoElement   ' 创造新的undo元素
    Dim c As Integer

    ' 删除全部Redo元素
    For c = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c

    ' 给新元素附值
    newElement.SelStart = rtfText.SelStart
    newElement.TextLen = Len(rtfText.Text)
    newElement.Text = rtfText.Text

    UndoStack.Add Item:=newElement    ' 添加到Undo堆栈

    EnableControls
End Sub

Private Sub rtfText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Ready!"    ' 当前状态提示
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then   ' 检查是否单击了鼠标右键。
        PopupMenu mnuEdit   ' 把文件菜单显示为一个弹出式菜单。
    End If
End Sub

