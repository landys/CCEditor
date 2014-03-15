Attribute VB_Name = "mdlEdit"
'**************************************************************************************'
'*     ��ģ����Ҫʵ�����ݵĸ��ơ�ճ�������С�ɾ�����ܣ��ı��Ĳ��ҡ��������Լ��ಽ�������ܡ� *'
'*     ���� Clipboard ���и��ơ�ճ�������С�ɾ���Ƚϼ򵥣�ֱ�ӵ��� VB ���ṩ�ĺ����Ϳ����� *'
'* ��ʵ�ֲ������õ���˳�������ķ��������⻹���Ƿ����ִ�Сд�����������ѡ��滻����һ���� *'
'* ��һ������һ��ȫ���滻����ʵ������������������ʵ�ֵġ��ಽ�������Զ���Ԫ�����ͣ����ø�Ԫ *'
'* �صļ���������Undo��RedoԪ�أ��������ʵ�ֶಽ������                                  *'
'**************************************************************************************'

Option Explicit

Public gFindString As String    ' Ҫ���ҵ��ַ���
Public gFindCase As Boolean ' ��־�����Ƿ����ִ�Сд��True Ϊ���֣�False Ϊ������
Public gFindDirection As Integer    ' ������������0 Ϊ����������1 Ϊ��������
Public gCurPos As Integer   ' ��ǰ��������λ��
Public gFirstTime As Boolean    ' ��־�Ƿ��ǵ�һ��������True Ϊ��һ�Σ�False Ϊ���ǵ�һ��
Public gReplaceString As String ' �����滻���ҵ����ı����ı�

Public trapUndo As Boolean ' ��־ĳ�����Ƿ񱻸��ټ�¼
Public UndoStack As New Collection ' UndoԪ�صļ���
Public RedoStack As New Collection ' RedoԪ�صļ���

Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63

Public Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


' ����ѡ�ı�������������

Public Sub CopyText()
    On Error Resume Next    ' ��������
    
    Clipboard.Clear ' ���ճ����
    Clipboard.SetText FrmMain.ActiveForm.rtfText.SelText
End Sub

'����ѡ�ı�������������

Public Sub CutText()
    On Error Resume Next    ' ��������
    
    Clipboard.Clear
    Clipboard.SetText FrmMain.ActiveForm.rtfText.SelText
    FrmMain.ActiveForm.rtfText.SelText = vbNullString
End Sub

' �Ӽ�����ճ���ı�
Public Sub PasteText()
    On Error Resume Next    ' ��������
    
    FrmMain.ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

' ɾ����ѡ���ı�

Public Sub DeleteText()
    FrmMain.ActiveForm.rtfText.SelText = vbNullString
End Sub

' ����ָ�����ַ�������
' ���ز��ҵ����ַ�����λ��

Private Function Find() As Integer
    Dim intStart As Integer ' �����ַ��������
    Dim strFindString As String ' Ҫ���ҵ��ַ���
    Dim strSourceString As String   ' �����ַ�����Χ
    Dim intPos As Integer   ' ���ҵ����ַ�����λ��
    Dim intOffset As Integer    ' ƫ������0 �� 1���ڶ�β�����������ֹ�������ҵ�ͬһ�ַ���
    
    If (gCurPos = FrmMain.ActiveForm.rtfText.SelStart) Then ' ���ղ鵽һ���ַ������´β�����㽫����
        intOffset = 1
    Else
        intOffset = 0
    End If
    If gFirstTime Then  ' �����β���
        intOffset = 0
    End If
    intStart = FrmMain.ActiveForm.rtfText.SelStart + intOffset  ' ������������
    
    If gFindCase Then   ' �����ִ�Сд
        strFindString = gFindString
        strSourceString = FrmMain.ActiveForm.rtfText.Text
    Else
        strFindString = UCase(gFindString)  ' �����ָ�ȫ����Ϊ��д
        strSourceString = UCase(FrmMain.ActiveForm.rtfText.Text)
    End If
    
    ' ���� intPos Ϊ�㽫ָʾû��Ҫ���ҵ��ַ���
    If gFindDirection Then  ' �����ҷ�������
        intPos = InStr(intStart + 1, strSourceString, strFindString)    ' ����
    ElseIf intStart <> 0 Then   ' �����ϣ�������ʼ�㲻���ļ���ʼ��
        For intPos = intStart - 1 To 0 Step -1    ' ����
            If intPos = 0 Then
                Exit For
            End If
            If Mid(strSourceString, intPos, Len(strFindString)) = strFindString Then
                Exit For
            End If
        Next
    Else
        intPos = 0
    End If
    
    If intPos <> 0 Then ' ���ҵ�
        FrmMain.ActiveForm.rtfText.SelStart = intPos - 1
        FrmMain.ActiveForm.rtfText.SelLength = Len(strFindString)
    End If
    
    gCurPos = FrmMain.ActiveForm.ActiveControl.SelStart   ' ���õ�ǰ��������λ��
    gFirstTime = False  ' ��־����ͬ���Ĳ��Ҳ����ǵ�һ��
    
    Find = intPos   ' ���ҵ���λ�÷���
End Function

' ʵ�ֲ����ַ��������Բ��ҽ�����д���

Public Sub FindIt()
    Dim intPos As Integer   ' Ҫ���ҵ��ַ���λ��
    Dim strMsg As String    ' ��ʾ����Ϣ�ַ���
    
    Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
    
    intPos = Find   ' �õ����ҽ��
    
    If intPos <> 0 Then ' ���ҵ����ѽ������¸�����
        FrmMain.ActiveForm.SetFocus
    Else    ' ���Ҳ�������Ϣ����ʾ
        strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
        MsgBox strMsg, 0, App.Title
    End If
    
    Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
End Sub

' ʵ���滻�ַ��������Խ�����д���

Public Sub ReplaceIt()
    Dim intPos As Integer   ' Ҫ���ҵ��ַ���λ��
    Dim strMsg As String    ' ��ʾ����Ϣ�ַ���
    
    Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
    
    intPos = Find   ' �õ����ҽ��
    
    If intPos <> 0 Then ' ���ҵ����滻�����ѽ������¸�����
        FrmMain.ActiveForm.rtfText.SelRTF = gReplaceString  ' ���ҵ����ַ����滻
        If gFindDirection = 0 Then  ' ��������������Ҫ������λ��
            FrmMain.ActiveForm.rtfText.SelStart = gCurPos
        End If
        FrmMain.ActiveForm.SetFocus ' ���ڻ�ý���
    Else    ' ���Ҳ�������Ϣ����ʾ
        strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
        MsgBox strMsg, 0, App.Title
    End If
    
    Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
End Sub

' �滻���в��ҵ����ַ���

Public Sub ReplaceAll()
    Dim intPos As Integer   ' Ҫ���ҵ��ַ���λ��
    Dim strMsg As String    ' ��ʾ����Ϣ�ַ���
    Dim bolFlag As Boolean  ' ��־�Ƿ����ַ������滻����
    
    bolFlag = False ' ��ʼΪδ�滻
    Screen.MousePointer = 11    ' �������Ϊ�ȴ���״
    
    ' ѭ�����ϲ����滻
    Do
        intPos = Find   ' �õ����ҽ��
        
        If intPos = 0 Then  ' ��û�ҵ�
            If Not bolFlag Then ' ����û���ַ������滻������Ϣ����ʾ
                strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
                MsgBox strMsg, 0, App.Title
            End If
            Exit Do
        Else    ' �ҵ�
            FrmMain.ActiveForm.rtfText.SelRTF = gReplaceString ' ���ҵ����ַ����滻����Ҫ���ַ���
            If gFindDirection = 0 Then  ' ��������������Ҫ������λ��
                FrmMain.ActiveForm.rtfText.SelStart = gCurPos
            End If
            bolFlag = True  ' ��־���ַ������滻
        End If
    Loop
    
    If bolFlag Then ' ���滻�������ڻ�ý���
        FrmMain.ActiveForm.SetFocus
    End If
    
    Screen.MousePointer = 0 ' �ָ����Ϊͨ���ļ�ͷ��״
End Sub

' ʹ��ť��Ч����Ч

Public Sub EnableControls()
    FrmChild.mnuEditUndo.Enabled = UndoStack.Count > 1
    FrmChild.mnuEditRedo.Enabled = RedoStack.Count > 0
    'FrmChild.rtfText_SelChange
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
    Dim tmpParam As String
    Dim d As Long
    If Len(lParam1) > Len(lParam2) Then ' ����
        tmpParam = lParam1
        lParam1 = lParam2
        lParam2 = tmpParam
    End If
    d = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d, d)
End Function

' ����

Public Sub Undo()
    Dim chg As String
    Dim x As Long
    Dim DeleteFlag As Boolean ' ��־�����ı�����ɾ���ı�
    Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  ' ɾ�����ı�
            x = SendMessage(FrmChild.rtfText.hwnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            FrmChild.rtfText.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            FrmChild.rtfText.SelLength = objElement.TextLen - objElement2.TextLen
            FrmChild.rtfText.SelText = ""
            x = SendMessage(FrmChild.rtfText.hwnd, EM_HIDESELECTION, 0&, 0&)
        Else ' �������ı�
            Set objElement = UndoStack(UndoStack.Count - 1)
            Set objElement2 = UndoStack(UndoStack.Count)
            chg = Change(objElement.Text, objElement2.Text, _
                objElement2.SelStart + 1 + Abs(Len(objElement.Text) - Len(objElement2.Text)))
            FrmChild.rtfText.SelStart = objElement2.SelStart
            FrmChild.rtfText.SelLength = 0
            FrmChild.rtfText.SelText = chg
            FrmChild.rtfText.SelStart = objElement2.SelStart
            If Len(chg) > 1 And chg <> vbCrLf Then
                FrmChild.rtfText.SelLength = Len(chg)
            Else
                FrmChild.rtfText.SelStart = FrmChild.rtfText.SelStart + Len(chg)
            End If
        End If
        RedoStack.Add Item:=UndoStack(UndoStack.Count)
        UndoStack.Remove UndoStack.Count
    End If
    
    EnableControls
    trapUndo = True
    FrmChild.rtfText.SetFocus
End Sub

' ����

Public Sub Redo()
Dim chg As String
Dim DeleteFlag As Boolean ' ��־�����ı�����ɾ���ı�
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(FrmChild.rtfText.Text)
        If DeleteFlag Then  ' ɾ�����ı�
            Set objElement = RedoStack(RedoStack.Count)
            FrmChild.rtfText.SelStart = objElement.SelStart
            FrmChild.rtfText.SelLength = Len(FrmChild.rtfText.Text) - objElement.TextLen
            FrmChild.rtfText.SelText = ""
        Else ' �������ı�
            Set objElement = RedoStack(RedoStack.Count)
            chg = Change(FrmChild.rtfText.Text, objElement.Text, objElement.SelStart + 1)
            FrmChild.rtfText.SelStart = objElement.SelStart - Len(chg)
            FrmChild.rtfText.SelLength = 0
            FrmChild.rtfText.SelText = chg
            FrmChild.rtfText.SelStart = objElement.SelStart - Len(chg)
            If Len(chg) > 1 And chg <> vbCrLf Then
                FrmChild.rtfText.SelLength = Len(chg)
            Else
                FrmChild.rtfText.SelStart = FrmChild.rtfText.SelStart + Len(chg)
            End If
        End If
        UndoStack.Add Item:=objElement
        RedoStack.Remove RedoStack.Count
    End If
    
    EnableControls
    trapUndo = True
    FrmChild.rtfText.SetFocus
End Sub

