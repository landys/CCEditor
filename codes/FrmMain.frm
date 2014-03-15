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
'*     ��������ģ��ʵ�ָó���Ľ��沼�ֵȹ��ܡ���Ҫ����Ŀ¼�ļ�����ʵ�֡���������״ *'
'* ̬����ʵ�ֵȡ�                                                               *'
'*******************************************************************************'

Option Explicit

' �������Ӱ�����Ϣ�� Windows ����
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, _
    ByVal dwData As Long) As Long

' DirlistBox �ؼ���Ŀ¼�ı�

Private Sub dirEditor_Change()
    On Error GoTo Dealerror ' �����ļ��в����ڴ���
    
    Dim dPath As String  ' ����·��
    
    dPath = filEditor.Path  ' ����·�����Ա�ָ�
    filEditor.Path = dirEditor.Path ' ���� FilelistBox �ؼ�·��
    Exit Sub
Dealerror:  ' ��ʾ�����ָ�ԭ·��
    MsgBox ("Directory Not Available")
    filEditor.Path = dPath
    dirEditor.Path = dPath
End Sub

Private Sub dirEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Direction change on navigation"    ' ��ǰ״̬��ʾ
End Sub

' DrivelistBox �ؼ���Ŀ¼�ı�

Private Sub drvEditor_Change()
    On Error GoTo Dealerror ' ���ܷ��������ô��󣬱���δ������̵�
    
    FrmMain.StatusBar1.Panels(1).Text = "Drive change on navigation."    ' ��ǰ״̬��ʾ
    
    Dim dPath As String  ' ����·��
    
    dPath = dirEditor.Path  ' ����·�����Ա�ָ�
    dirEditor.Path = drvEditor.Drive ' ���� DirlistBox �ؼ�������
    Exit Sub
Dealerror:  ' ��ʾ�����ָ�ԭ����
    MsgBox ("Drive Not Available")
    dirEditor.Path = dPath
    drvEditor.Drive = dPath
End Sub

' ˫�����ļ�

Private Sub filEditor_DblClick()
    On Error Resume Next
    
    Dim strFileName As String
    
    strFileName = filEditor.Path + "\" + filEditor.FileName
    NewFile  ' �����ĵ�����
    Me.ActiveForm.rtfText.LoadFile strFileName  ' �����ļ�
    Me.ActiveForm.Caption = strFileName
    Me.ActiveForm.m_Modified = False
    
    UpdateFileMenu (strFileName)    ' ���������ʾ�ļ��б��ú�����ģ�� mdlFile ��
End Sub


Private Sub filEditor_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Files change on navigation."    ' ��ǰ״̬��ʾ
End Sub

' ������س�ʼ��

Private Sub MDIForm_Load()
    ' ��ʼ���ֱ���
    gFindString = ""
    gReplaceString = ""
    gFindCase = False
    gFindDirection = 1
    gCurPos = 0
    gCountFrmChild = 0   ' ��ֵȫ�ֱ������Ӵ������
    
    ' ��ע������޴˳�����Ϣ����д��
    If GetSetting(App.Title, "Recent Files", "Maximum Of Recent Files") = Empty Then
        SaveSetting App.Title, "Recent Files", "Maximum Of Recent Files", 6
    End If
    ' ��ע����ȡ������ĵ����ֵ
    gMaxRecentFiles = GetSetting(App.Title, "Recent Files", "Maximum Of Recent Files")
    NewFile ' �½��ļ����ú�����ģ�� mdlFile ��
    
    GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
    FrmMain.StatusBar1.Panels(1).Text = "For Help, Press F1"    ' ��ǰ״̬��ʾ
End Sub

Private Sub MDIForm_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "For Help, Press F1"    ' ��ǰ״̬��ʾ
End Sub

Private Sub MDIForm_Resize()
    PicNavigation_Resize    ' ����Ŀ¼�ļ�������ͼƬ�������С
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Unload frmFind  ' �رղ�����ģʽ�Ի���
End Sub

Private Sub mnuFileExit_Click()
    Unload Me   ' �˳�
End Sub

' �������ļ�

Private Sub mnuFileNew_Click()
    NewFile ' �ú�����ģ�� mdlFile ��
End Sub

' ���ļ�

Private Sub mnuFileOpen_Click()
    OpenFile ' �ú�����ģ�� mdlFile ��
End Sub

' ������ļ��˵�

Private Sub mnuFileRec_Click(index As Integer)
    On Error GoTo Dealerror
    
    Dim strFileName As String   ' �ļ���
    Dim strKey As String    ' ��ע����еĹؼ�����
    
    If index = 0 Then ' ��� Empty �˵�ֱ�ӽ�������
        Exit Sub
    End If
    
    strKey = "Recent File" & index  ' ��Windows ע����е�Ӧ�ó�����Ŀ����ע���������ֵ
    strFileName = GetSetting(App.Title, "Recent Files", strKey)
    NewFile  ' �����ĵ����壬�ú�����ģ�� mdlFile ��
    Me.ActiveForm.rtfText.LoadFile strFileName
    Me.ActiveForm.Caption = strFileName
    UpdateFileMenu (strFileName)
    
    Me.ActiveForm.m_Modified = False ' ���ñ�־Ϊδ�޸�
    Exit Sub
    
Dealerror:  ' �ļ������ڴ�������ʾ��Ϣ����ɾ����Ӧע�����
    Unload Me.ActiveForm   ' ���ڴ��ļ�ǰ�½��˴��壬�ڴ���ɾ��
    MsgBox "The File Not Exist"
    DeleteRecentFile index  ' ɾ����Ӧע�����ú�����ģ�� mdlFile ��
    GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
End Sub

Private Sub mnuHelpAboutCCEditor_Click()
    frmAbout.Show 1 ' ��ʾ About �Ի���
End Sub

Private Sub mnuHelpTopic_Click()
    WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' ��ʾ������Ϣ
End Sub

' �����������ļ��б�˵�

Private Sub mnuOptionsClearRec_Click()
    DeleteAllRecentFiles    ' �����������ļ��б�˵����ú�����ģ�� mdlFile ��
    GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
End Sub

' ������ĵ��б������Ŀ���ã����ܳ���9

Private Sub mnuOptionsNumber_Click()
    On Error Resume Next
    
    Dim intN As Integer ' ��������ô�С
    Dim strN As String  ' ������Ի���������ı�
    Dim intCount As Integer ' ����������
    Dim i As Integer    ' For ѭ����������
    
    intCount = 0
    ' ����Ի���
    strN = InputBox("Input maximum of recent files showed on menu. The number should be 0--9.")
    intN = Asc(strN) - Asc("0") ' �õ����ִ�С
    If Len(strN) > 1 Then   ' ���������ַ����϶�Ϊ��
        intN = -1
    End If
    
    While intN < 0 Or intN > 10 ' ����
        intCount = intCount + 1 ' ������������ 1
        If (intCount = 3) Then  ' �����������������Ϣ��ʾ
            MsgBox "Input wrong for three times! Maximum of recent files not changed"
            Exit Sub
        End If
        ' ��ʾ������
        MsgBox "Input wrong! Please input again."
        intN = InputBox("Input maximum of recent files showed on menu. The number should be 0--9.")
    Wend
    If gMaxRecentFiles = intN Then
        Exit Sub
    ElseIf gMaxRecentFiles > intN Then  ' ��Ҫ���õ�ֵ��ԭ����С��ɾ������ע����������������ļ��б�
        For i = gMaxRecentFiles To intN + 1 Step -1
            DeleteRecentFile i  ' ɾ����Ӧע�����ú�����ģ�� mdlFile ��
        Next i
        GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
    End If
    
    gMaxRecentFiles = intN  ' ����������ļ��������ֵ
    ' ����ע�����
    SaveSetting App.Title, "Recent Files", "Maximum Of Recent Files", gMaxRecentFiles
    Exit Sub
End Sub

' ��ʾ�����س�����ߵ�Ŀ¼�ļ�����

Private Sub mnuViewNavigation_Click()
    mnuViewNavigation.Checked = Not mnuViewNavigation.Checked
    Me.PicNavigation.Visible = Not Me.PicNavigation.Visible
End Sub

' ��ʾ������״̬��

Private Sub mnuViewStatusbar_Click()
    mnuViewSTatusbar.Checked = Not mnuViewSTatusbar.Checked
    Me.StatusBar1.Visible = Not Me.StatusBar1.Visible
End Sub

' ��ʾ�����ع�����

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    Me.Toolbar1.Visible = Not Me.Toolbar1.Visible
End Sub

Private Sub PicNavigation_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Look on navigation."    ' ��ǰ״̬��ʾ
End Sub

' ����Ŀ¼�ļ�������ͼƬ�������С
    
Private Sub PicNavigation_Resize()
    On Error Resume Next
    
    drvEditor.Move 0, 0
    dirEditor.Move 0, (drvEditor.Height + 20), drvEditor.Width, (PicNavigation.Height - (drvEditor.Height + 20)) / 3
    filEditor.Move 0, drvEditor.Height + dirEditor.Height + 40, drvEditor.Width, PicNavigation.Height - dirEditor.Height - drvEditor.Height - 30
End Sub

Private Sub StatusBar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Cursor On Statusbar"    ' ��ǰ״̬��ʾ
End Sub

' ʵ�ֹ�����

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    FrmMain.StatusBar1.Panels(1).Text = Button.ToolTipText  ' ��ǰ״̬��ʾ
    Select Case Button.Key
        Case "new"
            NewFile ' �������ļ����ú�����ģ�� mdlFile ��
        Case "open"
            OpenFile    ' ���ļ����ú�����ģ�� mdlFile ��
        Case "save"
            SaveFile    ' �����ļ����ú�����ģ�� mdlFile ��
        Case "undo"
            Undo    ' ����
        Case "redo"
            Redo    ' ����
        Case "cut"
            CutText    ' ���У��ú�����ģ�� mdlEdit ��
        Case "copy"
            CopyText    ' �������ú�����ģ�� mdlEdit ��
        Case "paste"
            PasteText   ' ճ�����ú�����ģ�� mdlEdit ��
        Case "delete"
            DeleteText  ' ɾ�����ú�����ģ�� mdlEdit ��
        Case "help"
            WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' ��ʾ������Ϣ
    End Select
End Sub

Private Sub Toolbar1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Cursor On Toolbar"    ' ��ǰ״̬��ʾ
End Sub

' ���������幤��

Private Sub mnuOptionsToolsClipboard_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\clipbrd.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ��������������

Private Sub mnuOptionsToolsCalculator_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\calc.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ������ͼ�幤��

Private Sub mnuOptionsToolsPaint_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mspaint.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ������Ϸɨ��

Private Sub mnuOptionsGamesMiner_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\winmine.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ������Ϸ���Ľ���

Private Sub mnuOptionsGamesHearts_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mshearts.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"

End Sub

' ������Ϸֽ��

Private Sub mnuOptionsGamesSol_Click()

    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\sol.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ���� IE ���������

Private Sub mnuOptionsToolsIE_Click()
    On Error GoTo Dealerror
    
    Shell "C:\Program Files\Internet Explorer\IEXPLORE.EXE", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub
