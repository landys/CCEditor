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
'*     ���Ӵ���ģ��ʵ�ָó������Ҫ���ܡ�����Ҫ��Ϊ�ؼ� RichTextBox ��ʵ�֣������� *'
'* �롢���ļ���д�ļ����˵��ϵ�������Ӧ������Ƚ���Ҫ�Ĺ���ʵ���У���������״̬���� *'
'* ʵ�֣�������ļ��˵��б��ʵ�֣����½���򿪵��ļ����б༭�����롢�޸ġ�ɾ���� *'
'* ��ʵ�֣����а塢���ҡ��滻���ܵ�ʵ�֣��ಽ�������ܵ�ʵ�֣��ļ��е����֡��ַ�Ӧ�� *'
'* �ò�ͬ����ɫ��ʾ�����õ�ʵ�֣����߰������ܵ�ʵ�֣��ṩ�������߹��ܵ�ʵ�ֵȡ�     *'
'*******************************************************************************'

Option Explicit

Public m_Modified As Boolean    ' ��־�ļ��Ƿ����޸�
Private m_CharColor As Long  ' �ַ���ɫ
Private m_NumColor As Long   ' �����ַ���ɫ
Private m_TextChange As Boolean  ' �� RichTextBox ���ָı�ʱʹ��
                                 ' ��Ҫ�����ַ������Բ�ͬ��ɫ��ʾ�Ĺ���ʵ��

' �������Ӱ�����Ϣ�� Windows ����
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" _
    (ByVal hwnd As Long, ByVal lpHelpFile As String, ByVal wCommand As Long, _
    ByVal dwData As Long) As Long


' ���ڱ�����

Private Sub Form_Activate()
    GetRecentFiles  ' �ú�����ģ�� mdlFile �У����²˵���������ĵ��б�
End Sub

' ������س�ʼ��

Private Sub Form_Load()
    gCountFrmChild = gCountFrmChild + 1   ' ��ֵȫ�ֱ������Ӵ�������� 1
    Form_Resize ' ʵ�ִ����ڵĸ��ؼ�����
    m_CharColor = 0  ' �ַ���ɫ
    m_NumColor = 0   ' �����ַ���ɫ
End Sub

' �������ǰ��ʼ��

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim intResult As Integer
    If m_Modified Then
    '�����ļ����޸Ĺ�����ʾ�Ƿ񱣴�
        intResult = MsgBox("Save modified file to " & Me.Caption, _
            vbYesNoCancel + vbQuestion, "CCEditor")
        If intResult = vbCancel Then
            Cancel = 1
        ElseIf intResult = vbYes Then
            mnuFileSave_Click   '�����ļ����˳�
            Cancel = 0
        Else
            Cancel = 0  '�������˳�
        End If
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    ' �Ӵ��ڴ�С�仯ʱ���ı�����ű仯
    rtfText.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
    rtfText.RightMargin = rtfText.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gCountFrmChild = gCountFrmChild - 1 ' ��ֵȫ�ֱ������Ӵ�������� 1
End Sub

Private Sub mnuEditCopy_Click()
    CopyText    ' �������ú�����ģ�� mdlEdit ��
End Sub

Private Sub mnuEditCut_Click()
    CutText    ' ���У��ú�����ģ�� mdlEdit ��
End Sub

Private Sub mnuEditDelete_Click()
    DeleteText  ' ɾ�����ú�����ģ�� mdlEdit ��
End Sub

Private Sub mnuEditPaste_Click()
    PasteText   ' ճ�����ú�����ģ�� mdlEdit ��
End Sub

Private Sub mnuEditRedo_Click()
    Redo    ' ����
End Sub

' ѡ����������

Private Sub mnuEditSelectAll_Click()
    rtfText.SelStart = 0
    rtfText.SelLength = Len(rtfText.TextRTF)
End Sub

Private Sub mnuEditUndo_Click()
    Undo    ' ����
End Sub

Private Sub mnuFileClose_Click()
    Unload Me   ' �ر��ļ�
End Sub

Private Sub mnuFileExit_Click()
    Unload FrmMain  ' �رճ���
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
    rtfText.LoadFile strFileName
    Me.Caption = strFileName
    UpdateFileMenu (strFileName)
    Me.SetFocus
    
    m_Modified = False ' ���ñ�־Ϊδ�޸�
    Exit Sub
    
Dealerror:  ' �ļ������ڴ�������ʾ��Ϣ����ɾ����Ӧע�����
    Unload Me   ' ���ڴ��ļ�ǰ�½��˴��壬�ڴ���ɾ��
    MsgBox "The File Not Exist"
    DeleteRecentFile index  ' ɾ����Ӧע�����ú�����ģ�� mdlFile ��
    GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
End Sub

Private Sub mnuFileSaveAs_Click()
    SaveFileAs  ' ���Ϊ���ú�����ģ�� mdlFile ��
End Sub

Private Sub mnuHelpAboutCCEditor_Click()
    frmAbout.Show 1 ' ��ʾ About �Ի���
End Sub

Private Sub mnuHelpTopic_Click()
    WinHelp Me.hwnd, App.Path & "\Hlp\CCEditor.hlp", &H101, CLng(0) ' ��ʾ������Ϣ
End Sub

' ����Ϊ�м����

Private Sub mnuOptionsCenter_Click()
    
    mnuOptionsLeft.Checked = False
    mnuOptionsRight.Checked = False
    mnuOptionsCenter.Checked = True
    rtfText.SelAlignment = rtfCenter
End Sub

' �����������ļ��б�˵�

Private Sub mnuOptionsClearRec_Click()
    DeleteAllRecentFiles    ' �����������ļ��б�˵����ú�����ģ�� mdlFile ��
    GetRecentFiles  ' ����������ļ��б�˵����ú�����ģ�� mdlFile ��
End Sub

' �ַ���ɫ����

Private Sub mnuOptionsColorChar_Click()
    On Error Resume Next
    
    Dim lgSelStart As Long
    Dim lgSelLength As Long
    Dim i As Long
    
    m_TextChange = False    ' ��־�������� rtfText_Change ���̱�����ʱ���������޹ش���
    
    ' ��ɫͨ�öԻ����ʼ��
    FrmMain.dlgCommonDialog.Color = m_CharColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    ' ����ͨ��һ�����ַ��ж�ʵ��
    If m_CharColor <> FrmMain.dlgCommonDialog.Color Then    ' ��ɫ�Ѹı�
        m_CharColor = FrmMain.dlgCommonDialog.Color
        If rtfText.SelText <> "" Then   ' ���ı�ѡ��
            lgSelStart = rtfText.SelStart   ' ����ѡ���ı���λ�ã��Ա�ָ�
            lgSelLength = rtfText.SelLength   ' ����ѡ���ı��ĳ��ȣ��Ա�ָ�
            For i = 0 To lgSelLength - 1    ' ���ַ��Ƚϣ��ı���ɫ
                rtfText.SelStart = lgSelStart + i
                rtfText.SelLength = 1
                If Not (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then
                    ' �������ַ�
                    rtfText.SelColor = m_CharColor
                End If
            Next
            rtfText.SelStart = lgSelStart ' �ָ�ѡ���ı���λ��
            rtfText.SelLength = lgSelLength   ' �ָ�ѡ���ı��ĳ���
        End If
    End If
    
    m_TextChange = True ' �ָ���־
End Sub

' ������ɫ����

Private Sub mnuOptionsColorNum_Click()
    On Error Resume Next
    
    Dim lgSelStart As Long
    Dim lgSelLength As Long
    Dim i As Long
    
    m_TextChange = False    ' ��־�������� rtfText_Change ���̱�����ʱ���������޹ش���
    
    ' ��ɫͨ�öԻ����ʼ��
    FrmMain.dlgCommonDialog.Color = m_NumColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    ' ����ͨ��һ�����ַ��ж�ʵ��
    If m_NumColor <> FrmMain.dlgCommonDialog.Color Then    ' ��ɫ�Ѹı�
        m_NumColor = FrmMain.dlgCommonDialog.Color
        If rtfText.SelText <> "" Then   ' ���ı�ѡ��
            lgSelStart = rtfText.SelStart   ' ����ѡ���ı���λ�ã��Ա�ָ�
            lgSelLength = rtfText.SelLength   ' ����ѡ���ı��ĳ��ȣ��Ա�ָ�
            For i = 0 To lgSelLength - 1    ' ���ַ��Ƚϣ��ı���ɫ
                rtfText.SelStart = lgSelStart + i
                rtfText.SelLength = 1
                If (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then ' �������ַ�
                    rtfText.SelColor = m_NumColor
                End If
            Next
            rtfText.SelStart = lgSelStart ' �ָ�ѡ���ı���λ��
            rtfText.SelLength = lgSelLength   ' �ָ�ѡ���ı��ĳ���
        End If
    End If
    
    m_TextChange = True ' �ָ���־
End Sub

' �ı���ɫ����

Private Sub mnuOptionsColorText_Click()
    m_TextChange = False    ' ��־�������� rtfText_Change ���̱�����ʱ���������޹ش���
    
    ' ��ɫͨ�öԻ����ʼ��
    FrmMain.dlgCommonDialog.Color = m_CharColor
    FrmMain.dlgCommonDialog.ShowColor
    FrmMain.dlgCommonDialog.CancelError = True
    
    m_CharColor = FrmMain.dlgCommonDialog.Color
    m_NumColor = m_CharColor
    rtfText.SelColor = m_CharColor  ' �ı���ɫ����
    
    m_TextChange = True ' �ָ���־
End Sub

' ��������

Private Sub mnuOptionsFont_Click()
    On Error Resume Next
    
    m_TextChange = False    ' ��־�������� rtfText_Change ���̱�����ʱ���������޹ش���
    
    ' ��ͨ�öԻ�����������
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
    
    m_TextChange = True ' �ָ���־
End Sub

' ������Ϸ���Ľ���

Private Sub mnuOptionsGamesHearts_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\mshearts.exe", 1
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

' ������Ϸֽ��

Private Sub mnuOptionsGamesSol_Click()

    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\sol.exe", 1
    Exit Sub
    
Dealerror:
    MsgBox "Cannot Find the Execute File"
End Sub

' ����Ϊ�ı������
    
Private Sub mnuOptionsLeft_Click()
    mnuOptionsLeft.Checked = True
    mnuOptionsRight.Checked = False
    mnuOptionsCenter.Checked = False
    rtfText.SelAlignment = rtfLeft
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

' ����Ϊ�ı��Ҷ���

Private Sub mnuOptionsRight_Click()
    mnuOptionsLeft.Checked = False
    mnuOptionsRight.Checked = True
    mnuOptionsCenter.Checked = False
    rtfText.SelAlignment = rtfRight
End Sub

' ���������幤��

Private Sub mnuOptionsToolsClipboard_Click()
    On Error GoTo Dealerror
    
    Shell "C:\WINDOWS\system32\clipbrd.exe", 1
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

' ����

Private Sub mnuSearchFind_Click()
    ' ���ҵĳ�ʼֵΪ�ϴβ��ҵ�ֵ�����ļ�ѡ�е��ַ�
    If Me.rtfText.SelText <> "" Then
        frmFind.txtFind.Text = Me.rtfText.SelText
    Else
        frmFind.txtFind.Text = gFindString
    End If
    
    gFirstTime = True  ' ȫ�ֱ����������״β���
    
    If (gFindCase) Then ' �����Ƿ����ִ�Сд
        frmFind.chkCase = 1
    End If
    
    frmFind.Show 0  ' ��ʾ���ҶԻ���
End Sub

' ������һ��
Private Sub mnuSearchFindNext_Click()
    If Len(gFindString) > 0 Then    ' �����ҵ��ı���Ϊ��
        gFirstTime = False
        FindIt  ' ʵ�ֲ��ң��ú�����ģ�� mdlEdit �ж���
    Else    ' ������Ϊ��
        mnuSearchFind_Click ' ������Ϊ��
    End If
End Sub

' �滻

Private Sub mnuSearchReplace_Click()
    frmFind.txtReplace.Text = gReplaceString
    mnuSearchFind_Click
End Sub

' ��ʾ�����س�����ߵ�Ŀ¼�ļ�����

Private Sub mnuViewNavigation_Click()
    mnuViewNavigation.Checked = Not mnuViewNavigation.Checked
    FrmMain.PicNavigation.Visible = Not FrmMain.PicNavigation.Visible
End Sub

' ��ʾ������״̬��

Private Sub mnuViewStatusbar_Click()
    mnuViewSTatusbar.Checked = Not mnuViewSTatusbar.Checked
    FrmMain.StatusBar1.Visible = Not FrmMain.StatusBar1.Visible
End Sub

' ��ʾ�����ع�����

Private Sub mnuViewToolbar_Click()
    mnuViewToolbar.Checked = Not mnuViewToolbar.Checked
    FrmMain.Toolbar1.Visible = Not FrmMain.Toolbar1.Visible
End Sub

' ������MDI��С�����Ӵ���ͼ���������
    
Private Sub mnuWindowArrangeIcons_Click()
    FrmMain.Arrange vbArrangeIcons
End Sub

' ������MDI����С�����Ӵ������ŷ�

Private Sub mnuWindowCascade_Click()
    FrmMain.Arrange vbCascade
End Sub

' ������MDI����С�����Ӵ���ˮƽ�ŷ�

Private Sub mnuWindowTileHor_Click()
    FrmMain.Arrange vbTileHorizontal
End Sub

' ������MDI����С�����Ӵ��崹ֱ�ŷ�

Private Sub mnuWindowTileVer_Click()
    FrmMain.Arrange vbTileVertical
End Sub

' �����ļ�

Private Sub mnuFileSave_Click()
    SaveFile     ' �ú�����ģ�� mdlFile ��
End Sub

Private Sub rtfText_Change()
    ' ����ı��б仯���������޸ı�־
    m_Modified = True

    ' �ַ�������ж��������ֻ��Ƿ����֣����Բ�ͬ��ɫ��ʾ
    If m_TextChange And m_CharColor <> 0 And m_NumColor <> 0 Then
        rtfText.SelStart = rtfText.SelStart - 1 ' ѡ������ַ����ٸ�Ϊ��Ӧ��ɫ
        rtfText.SelLength = 1
        If (rtfText.SelText <= "9" And rtfText.SelText >= "0") Then
            rtfText.SelColor = m_NumColor
        Else
            rtfText.SelColor = m_CharColor
        End If
        
        rtfText.SelStart = rtfText.SelStart + 1 '�ָ����
        rtfText.SelLength = 0   ' ȡ��ѡ��
    End If
    
    If Not trapUndo Then  ' �������ٱ���ֹ
        Exit Sub
    End If

    Dim newElement As New UndoElement   ' �����µ�undoԪ��
    Dim c As Integer

    ' ɾ��ȫ��RedoԪ��
    For c = 1 To RedoStack.Count
        RedoStack.Remove 1
    Next c

    ' ����Ԫ�ظ�ֵ
    newElement.SelStart = rtfText.SelStart
    newElement.TextLen = Len(rtfText.Text)
    newElement.Text = rtfText.Text

    UndoStack.Add Item:=newElement    ' ��ӵ�Undo��ջ

    EnableControls
End Sub

Private Sub rtfText_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    FrmMain.StatusBar1.Panels(1).Text = "Ready!"    ' ��ǰ״̬��ʾ
End Sub

Private Sub rtfText_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then   ' ����Ƿ񵥻�������Ҽ���
        PopupMenu mnuEdit   ' ���ļ��˵���ʾΪһ������ʽ�˵���
    End If
End Sub

