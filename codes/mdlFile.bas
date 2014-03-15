Attribute VB_Name = "mdlFile"
'**********************************************************************************'
'*     ��ģ���й����ļ������������ļ��½����򿪡����漰�˵�������ļ��б�ĸ����й��ļ� *'
'* ���������������� RichTextBox �ؼ������Դ��ļ�ʵ������ͨ���½�һ�����壬Ȼ����ļ� *'
'* ���뵽�ÿؼ���ʵ�֣�ͬʱ�ÿ���Ҳ�޶��˱�����ֻ��֧�� .txt �� .rtf ��ʽʽ���ļ���    *'
'*     ����ļ��б���µ���Ϣ�洢��ϵͳע����У����Ը���صĺ�����Ҫ��ע�����в����� *'
'* ������д��ɾ��ע�����������ڱ�������ö��ĵ��ṹ��֧��ͬʱ�򿪶���ļ����� VB �� *'
'* �����ڵĲ˵����Ӵ��ڵĲ˵��Ƿֿ��ģ���Ȼ���Ӵ���ʱ����˵���������ڵĲ˵��滻���� *'
'* ���ԶԲ˵��е�������ĵ��ĸ���Ҫͬʱ�����ߵĲ˵����и���                         *'
'**********************************************************************************'

Option Explicit

Public gMaxRecentFiles As Integer    ' ������ĵ�����������ֵ0��9
Public gFileNameArray As Variant ' �洢�ļ���������
Private gIndex As Integer   ' ������ĵ�������
Public gCountFrmChild As Integer    ' ���������Ӵ������


' ���򿪵��ļ��Ƿ��Ѿ�������������б���
' �����ڣ����� True �����򣬷��� False
' ��������Ϊ�򿪵��ļ���

Private Function OnRecentFilesList(FileName As String) As Integer
    On Error GoTo Dealerror ' ����gFileNameArray ����δ��ֵ����������
    
    Dim i As Integer
    
    For i = 1 To UBound(gFileNameArray, 1)  ' �����ļ��б���
        If gFileNameArray(i, 1) = FileName Then  ' �ļ��б������Ѵ��ڴ��ļ���
            gIndex = i  ' ������ļ����ļ��б������кţ��Ա������������н����Ƶ���ǰ
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    
    OnRecentFilesList = False
    Exit Function

Dealerror:  ' ��������ת����
    OnRecentFilesList = False
End Function


' �����ļ���������ĵ��б�
' ��������Ϊ�򿪵��ļ���

Public Sub UpdateFileMenu(FileName As String)
    Dim intRetVal As Integer
     
    intRetVal = OnRecentFilesList(FileName)
    If Not intRetVal Then   ' �жϴ򿪵��ļ����Ƿ��Ѿ��ڡ��ļ����˵��ؼ�������
        WriteRecentFiles FileName, gMaxRecentFiles ' ���򿪵��ļ���д��ע���
    ElseIf gIndex > 1 Then
        WriteRecentFiles FileName, gIndex
    End If
    GetRecentFiles  ' ���¡��ļ����˵��ؼ����鼰����򿪵��ļ��б�
End Sub

' ��ע����ж�������ĵ��б��¼�����µ��˵���

Public Sub GetRecentFiles()
    Dim i As Integer    ' ����
    
    If GetSetting(App.Title, "Recent Files", "Recent File1") = Empty Then
        ' ���ע�������Ӧ��Ŀ�Ƿ�Ϊ��
        FrmMain.mnuFileRec(0).Visible = True    ' ��ʾ�˵��пյ�������ĵ��б�
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileRec(0).Visible = True
        End If
        For i = 1 To gMaxRecentFiles
            FrmMain.mnuFileRec(i).Visible = False   ' �����Ѿ���ʾ��������ĵ�
            If gCountFrmChild > 0 Then
                FrmMain.ActiveForm.mnuFileRec(i).Visible = False
            End If
        Next i
        FrmMain.mnuFileSeperator3.Visible = False  ' �����ļ��б������˳��˵�֮��ķָ���
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileSeperator3.Visible = False
        End If
        
        gFileNameArray = Empty
        Exit Sub
    End If
    
    gFileNameArray = GetAllSettings(App.Title, "Recent Files")  ' ��ע����еõ�������ĵ���Ϣ
    For i = 1 To UBound(gFileNameArray, 1)   ' ��ȡ��������±�
        ' ����ļ�����·��
        FrmMain.mnuFileRec(i).Caption = "&" & Format$(i) & " " & gFileNameArray(i, 1)
        FrmMain.mnuFileRec(i).Visible = True   ' ��ʾע����д��ڵ�������ĵ�
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileRec(i).Caption = "&" & Format$(i) & " " & gFileNameArray(i, 1)
            FrmMain.ActiveForm.mnuFileRec(i).Visible = True
        End If
    Next i
    FrmMain.mnuFileRec(0).Visible = False
    FrmMain.mnuFileSeperator3.Visible = True  ' ��ʾ�ļ��б������˳��˵�֮��ķָ���
    If gCountFrmChild > 0 Then
        FrmMain.ActiveForm.mnuFileRec(0).Visible = False
        FrmMain.ActiveForm.mnuFileSeperator3.Visible = True
    End If
End Sub
  
' ��������ĵ���Ϣд��ע���
' ��һ������Ϊ�ļ������ڶ���Ϊ����������ĵ��е����к�

Private Sub WriteRecentFiles(FileName As String, index As Integer)
    Dim i As Integer    ' ����
    Dim strFileName As String   ' �ļ���
    Dim strKey As String    ' ��ע����еĹؼ�����
    
    ' ���ļ�Recent File1 ���Ƹ�Recent File2...��ʹ�ļ��б���˳������
    For i = index - 1 To 1 Step -1
        strKey = "Recent File" & i
        ' ��Windows ע����е�Ӧ�ó�����Ŀ����ע���������ֵ
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            strKey = "Recent File" & (i + 1)
            SaveSetting App.Title, "Recent Files", strKey, strFileName 'д��ע���
        End If
    Next i
    ' �����´򿪵��ļ�д�����ʹ���ļ��б�ĵ�һ��
    SaveSetting App.Title, "Recent Files", "Recent File1", FileName
End Sub

' ��ָ��������ĵ���Ϣ��ע���ɾ��
' ����Ϊ����������ĵ��е����к�

Public Sub DeleteRecentFile(index As Integer)
    Dim i As Integer    ' ����
    Dim strFileName As String   ' �ļ���
    Dim strKey As String    ' ��ע����еĹؼ�����
    
    If index > gMaxRecentFiles Then ' �����������ע����д���
        Exit Sub
    End If
    
    ' ���ļ��б���˳�����Ƶ��� index ����ɾ�����һ��
    For i = index + 1 To gMaxRecentFiles
        strKey = "Recent File" & i
        ' ��Windows ע����е�Ӧ�ó�����Ŀ����ע���������ֵ
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            strKey = "Recent File" & (i - 1)
            SaveSetting App.Title, "Recent Files", strKey, strFileName 'д��ע���
        Else
            Exit For
        End If
    Next i
    
    strKey = "recent file" & (i - 1)
    DeleteSetting App.Title, "Recent Files", strKey ' ɾ�����һ��
End Sub

' ������������ĵ���Ϣ��ע���ɾ��

Public Sub DeleteAllRecentFiles()
    Dim i As Integer    ' ����
    Dim strFileName As String   ' �ļ���
    Dim strKey As String    ' ��ע����еĹؼ�����

    ' ����ɾ��
    For i = 1 To gMaxRecentFiles
        strKey = "Recent File" & i
        ' ��Windows ע����е�Ӧ�ó�����Ŀ����ע���������ֵ
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            DeleteSetting App.Title, "Recent Files", strKey ' ɾ��ע�������Ϣ
        Else
            Exit For
        End If
    Next i
End Sub

' �������ļ�

Public Sub NewFile()
    Dim frmNew As New FrmChild  ' �����´������
    Static lgFileCount As Long  ' ��¼�½����ļ�����
    
    lgFileCount = lgFileCount + 1   ' ��������
    frmNew.Caption = "Noname" & lgFileCount ' ���´���ı�����Ϊ��Noname��+ ����
    frmNew.Show ' ��ʾ����
    frmNew.m_Modified = False ' ���ñ�־Ϊδ�޸�
End Sub

' �����ļ���ͨ���ȴ������ļ����壬�ٰ��ļ����������е� RichTextBox ��ʵ��

Public Sub OpenFile()
    On Error Resume Next    ' ���ÿ�����������
    
    Dim strFileName As String   ' �ļ���
    
    With FrmMain.dlgCommonDialog    ' ���ļ�ͨ�öԻ���ĳ�ʼ����
        .DialogTitle = "Open"
        .Filter = "Text (*.txt)|*.txt|RTFText (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
        If Err = 32755 Then ' �û���� Cancel ���˳�
            Exit Sub
        End If
        If Len(.FileName) = 0 Then  ' �����ļ���
            Exit Sub
        End If
        strFileName = .FileName
    End With
        
    Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
    
    NewFile  ' �½����ĵ�����
    FrmMain.ActiveForm.rtfText.LoadFile strFileName ' RichTextBox �м����ļ�
    FrmMain.ActiveForm.Caption = strFileName    ' �������Ϊ���ļ���
    FrmMain.ActiveForm.m_Modified = False ' ���ñ�־Ϊδ�޸�
    
    Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
    
    UpdateFileMenu (strFileName)    ' ����������ĵ��б�
End Sub

' �����ļ�

Public Sub SaveFile()
    On Error Resume Next    ' ���ÿ�����������
    
    Dim strFileName As String   'ȡ�ñ����ļ����������ļ�
    
    If Left(FrmMain.ActiveForm.Caption, 6) = "Noname" Then  ' ���ļ�δ���������������Ϊ
        SaveFileAs
    Else
        Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
        
        strFileName = FrmMain.ActiveForm.Caption    ' �������Ϊ���ļ���
        FrmMain.ActiveForm.rtfText.SaveFile strFileName ' �����ļ�
        FrmMain.ActiveForm.m_Modified = False ' ���ñ�־Ϊδ�޸�
        
        Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
    End If
End Sub

' �ļ����Ϊ

Public Sub SaveFileAs()
    On Error Resume Next    ' ���ÿ�����������
    
    Dim strFileName As String   ' �ļ���
    
    With FrmMain.dlgCommonDialog    ' ���Ϊͨ�öԻ���ܵ�ʵ��
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "Text (*.txt)|*.txt|RTFText (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowSave
        If Len(.FileName) = 0 Then  ' �����ļ���
            Exit Sub
        End If
        strFileName = .FileName
    End With
    
    Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
    
    FrmMain.ActiveForm.Caption = strFileName    ' �������Ϊ���ļ���
    FrmMain.ActiveForm.rtfText.SaveFile strFileName ' �����ļ�
    FrmMain.ActiveForm.m_Modified = False ' ���ñ�־Ϊδ�޸�
    
    Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
    
    UpdateFileMenu (strFileName)    ' ����������ĵ��б�
End Sub

