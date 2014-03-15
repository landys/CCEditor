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
'*     ���ҶԻ���ģ�飬�Ի����еĿؼ���Ӧ��ʵ�ֲ����滻���ܡ���Ҫ����Ϊ���ҡ�������һ�����滻 *'
'* ��һ����ȫ���滻�����⻹�����Ƿ����ִ�Сд�����������趨�ȡ�                             *'
'****************************************************************************************'

Option Explicit

' �����趨����λ�� Windows ������������ʹ�ô���ʼ��λ�����ϲ�
Private Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, _
    ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


' CheckBox���趨�Ƿ����ִ�С

Private Sub chkCase_Click()
    gFindCase = chkCase.Value   ' ��ȫ�ֱ�����ֵ
End Sub

' Cancel �˳�

Private Sub cmdCancel_Click()
    ' ��ȫ�ֱ�����ֵ�����´�ʹ��
    gFindString = txtFind.Text
    gReplaceString = txtReplace.Text
    gFindCase = chkCase.Value
    
    Unload frmFind  ' �˳�
End Sub

' ���Ұ�ť

Private Sub cmdFind_Click()
    gFindString = txtFind.Text   ' ��ȫ�ֱ�����ֵ
    cmdFind.Caption = "Find Next"
    FindIt  ' �˺�����ģ�� mdlEdit �У�ʵ�ֲ���
    
    gFirstTime = False  ' ��־�Ѳ��ҹ�
End Sub

' �滻��ť

Private Sub cmdReplace_Click()
    gFindString = txtFind.Text   ' ��ȫ�ֱ�����ֵ
    gReplaceString = txtReplace.Text
    ReplaceIt  ' �˺�����ģ�� mdlEdit �У�ʵ���滻
    
    gFirstTime = False  ' ��־�Ѳ��ҹ�
End Sub

' �滻ȫ����ť

Private Sub cmdReplaceAll_Click()
    gFindString = txtFind.Text   ' ��ȫ�ֱ�����ֵ
    gReplaceString = txtReplace.Text
    
    ReplaceAll  ' �˺�����ģ�� mdlEdit �ж��壬�滻����Ҫ���ҵ��ַ���
    
    gFirstTime = False  ' ��־�Ѳ��ҹ�
End Sub

' �������ʱ��ʼ��

Private Sub Form_Load()
    SetWindowPos Me.hwnd, -1, 510, 195, 345, 180, 0 ' �ô���������ʾ����������֮��
    
    cmdFind.Enabled = False ' Disable ���Ұ�ť
    cmdReplace.Enabled = False  ' Disable �滻��ť
    cmdReplaceAll.Enabled = False   ' Disable �滻ȫ����ť

    optDirection(gFindDirection).Value = 1  ' ��ȡ������������
End Sub

' OptionButton ������������

Private Sub optDirection_Click(index As Integer)
    gFindDirection = index  ' ������������
End Sub

' ����Ҫ���ҵ��ı����ݸı�

Private Sub txtFind_Change()
    gFirstTime = True   ' ��ȫ�ֱ�����ֵ
    
    cmdFind.Caption = "Find"    ' ���ò��Ұ�ť��ʾ��ϢΪ Find
    
    If txtFind.Text = "" Then   ' ���� TextBox Ϊ��ʱ��Disable ����ذ�ť
        cmdFind.Enabled = False
        cmdReplace.Enabled = False
        cmdReplaceAll.Enabled = False
    Else   ' ���� TextBox ��Ϊ��ʱ��Enable ����ذ�ť
        cmdFind.Enabled = True
        cmdReplace.Enabled = True
        cmdReplaceAll.Enabled = True
    End If
End Sub

