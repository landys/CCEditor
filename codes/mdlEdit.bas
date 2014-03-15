Attribute VB_Name = "mdlEdit"
'**************************************************************************************'
'*     本模块主要实现数据的复制、粘贴、剪切、删除功能，文本的查找、换功能以及多步撤销功能。 *'
'*     利用 Clipboard 进行复制、粘贴、剪切、删除比较简单，直接调用 VB 所提供的函数就可以了 *'
'* 而实现查找利用的是顺序搜索的方法，另外还有是否区分大小写、搜索方向等选项。替换包括一次替 *'
'* 换一个，或一次全部替换，这实际上是在搜索基础上实现的。多步撤销先自定义元素类型，再用该元 *'
'* 素的集合来保存Undo和Redo元素，这样便可实现多步撤销。                                  *'
'**************************************************************************************'

Option Explicit

Public gFindString As String    ' 要查找的字符串
Public gFindCase As Boolean ' 标志查找是否区分大小写，True 为区分，False 为不区分
Public gFindDirection As Integer    ' 设置搜索方向，0 为向上搜索，1 为向下搜索
Public gCurPos As Integer   ' 当前搜索到的位置
Public gFirstTime As Boolean    ' 标志是否是第一次搜索，True 为第一次，False 为不是第一次
Public gReplaceString As String ' 用来替换查找到的文本的文本

Public trapUndo As Boolean ' 标志某操作是否被跟踪记录
Public UndoStack As New Collection ' Undo元素的集合
Public RedoStack As New Collection ' Redo元素的集合

Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63

Public Declare Function SendMessage Lib "user32" Alias _
    "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
        ByVal wParam As Long, lParam As Long) As Long


' 把所选文本拷贝到剪贴板

Public Sub CopyText()
    On Error Resume Next    ' 跳过错误
    
    Clipboard.Clear ' 清除粘贴板
    Clipboard.SetText FrmMain.ActiveForm.rtfText.SelText
End Sub

'把所选文本剪贴到剪贴板

Public Sub CutText()
    On Error Resume Next    ' 跳过错误
    
    Clipboard.Clear
    Clipboard.SetText FrmMain.ActiveForm.rtfText.SelText
    FrmMain.ActiveForm.rtfText.SelText = vbNullString
End Sub

' 从剪贴板粘贴文本
Public Sub PasteText()
    On Error Resume Next    ' 跳过错误
    
    FrmMain.ActiveForm.rtfText.SelText = Clipboard.GetText
End Sub

' 删除所选的文本

Public Sub DeleteText()
    FrmMain.ActiveForm.rtfText.SelText = vbNullString
End Sub

' 查找指定的字符串函数
' 返回查找到的字符串的位置

Private Function Find() As Integer
    Dim intStart As Integer ' 查找字符串的起点
    Dim strFindString As String ' 要查找的字符串
    Dim strSourceString As String   ' 查找字符串范围
    Dim intPos As Integer   ' 查找到的字符串的位置
    Dim intOffset As Integer    ' 偏移量，0 或 1，在多次查找中用来防止反复查找到同一字符串
    
    If (gCurPos = FrmMain.ActiveForm.rtfText.SelStart) Then ' 若刚查到一个字符串，下次查找起点将下移
        intOffset = 1
    Else
        intOffset = 0
    End If
    If gFirstTime Then  ' 若初次查找
        intOffset = 0
    End If
    intStart = FrmMain.ActiveForm.rtfText.SelStart + intOffset  ' 查找起点的设置
    
    If gFindCase Then   ' 若区分大小写
        strFindString = gFindString
        strSourceString = FrmMain.ActiveForm.rtfText.Text
    Else
        strFindString = UCase(gFindString)  ' 不区分刚全部设为大写
        strSourceString = UCase(FrmMain.ActiveForm.rtfText.Text)
    End If
    
    ' 这里 intPos 为零将指示没有要查找的字符串
    If gFindDirection Then  ' 若查找方向向下
        intPos = InStr(intStart + 1, strSourceString, strFindString)    ' 查找
    ElseIf intStart <> 0 Then   ' 若向上，并且起始点不在文件起始点
        For intPos = intStart - 1 To 0 Step -1    ' 查找
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
    
    If intPos <> 0 Then ' 若找到
        FrmMain.ActiveForm.rtfText.SelStart = intPos - 1
        FrmMain.ActiveForm.rtfText.SelLength = Len(strFindString)
    End If
    
    gCurPos = FrmMain.ActiveForm.ActiveControl.SelStart   ' 设置当前搜索到的位置
    gFirstTime = False  ' 标志下面同样的查找不再是第一次
    
    Find = intPos   ' 查找到的位置返回
End Function

' 实现查找字符串，并对查找结果进行处理

Public Sub FindIt()
    Dim intPos As Integer   ' 要查找的字符串位置
    Dim strMsg As String    ' 显示的消息字符串
    
    Screen.MousePointer = 11    ' 设置鼠标为等待形状
    
    intPos = Find   ' 得到查找结果
    
    If intPos <> 0 Then ' 若找到，把焦点重新给窗口
        FrmMain.ActiveForm.SetFocus
    Else    ' 若找不到，消息框提示
        strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
        MsgBox strMsg, 0, App.Title
    End If
    
    Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
End Sub

' 实现替换字符串，并对结果进行处理

Public Sub ReplaceIt()
    Dim intPos As Integer   ' 要查找的字符串位置
    Dim strMsg As String    ' 显示的消息字符串
    
    Screen.MousePointer = 11    ' 设置鼠标为等待形状
    
    intPos = Find   ' 得到查找结果
    
    If intPos <> 0 Then ' 若找到，替换，并把焦点重新给窗口
        FrmMain.ActiveForm.rtfText.SelRTF = gReplaceString  ' 把找到的字符串替换
        If gFindDirection = 0 Then  ' 搜索方向若向上要重设光标位置
            FrmMain.ActiveForm.rtfText.SelStart = gCurPos
        End If
        FrmMain.ActiveForm.SetFocus ' 窗口获得焦点
    Else    ' 若找不到，消息框提示
        strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
        MsgBox strMsg, 0, App.Title
    End If
    
    Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
End Sub

' 替换所有查找到的字符串

Public Sub ReplaceAll()
    Dim intPos As Integer   ' 要查找的字符串位置
    Dim strMsg As String    ' 显示的消息字符串
    Dim bolFlag As Boolean  ' 标志是否有字符串被替换过了
    
    bolFlag = False ' 初始为未替换
    Screen.MousePointer = 11    ' 设置鼠标为等待形状
    
    ' 循环不断查找替换
    Do
        intPos = Find   ' 得到查找结果
        
        If intPos = 0 Then  ' 若没找到
            If Not bolFlag Then ' 并且没有字符串被替换过，消息框提示
                strMsg = "Cannot find" & Chr(34) & gFindString & Chr(34)
                MsgBox strMsg, 0, App.Title
            End If
            Exit Do
        Else    ' 找到
            FrmMain.ActiveForm.rtfText.SelRTF = gReplaceString ' 把找到的字符串替换成想要的字符串
            If gFindDirection = 0 Then  ' 搜索方向若向上要重设光标位置
                FrmMain.ActiveForm.rtfText.SelStart = gCurPos
            End If
            bolFlag = True  ' 标志有字符串被替换
        End If
    Loop
    
    If bolFlag Then ' 若替换过，窗口获得焦点
        FrmMain.ActiveForm.SetFocus
    End If
    
    Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
End Sub

' 使按钮有效或无效

Public Sub EnableControls()
    FrmChild.mnuEditUndo.Enabled = UndoStack.Count > 1
    FrmChild.mnuEditRedo.Enabled = RedoStack.Count > 0
    'FrmChild.rtfText_SelChange
End Sub

Public Function Change(ByVal lParam1 As String, ByVal lParam2 As String, startSearch As Long) As String
    Dim tmpParam As String
    Dim d As Long
    If Len(lParam1) > Len(lParam2) Then ' 交换
        tmpParam = lParam1
        lParam1 = lParam2
        lParam2 = tmpParam
    End If
    d = Len(lParam2) - Len(lParam1)
    Change = Mid(lParam2, startSearch - d, d)
End Function

' 撤销

Public Sub Undo()
    Dim chg As String
    Dim x As Long
    Dim DeleteFlag As Boolean ' 标志增加文本还是删除文本
    Dim objElement As Object, objElement2 As Object
    If UndoStack.Count > 1 And trapUndo Then
        trapUndo = False
        DeleteFlag = UndoStack(UndoStack.Count - 1).TextLen < UndoStack(UndoStack.Count).TextLen
        If DeleteFlag Then  ' 删除了文本
            x = SendMessage(FrmChild.rtfText.hwnd, EM_HIDESELECTION, 1&, 1&)
            Set objElement = UndoStack(UndoStack.Count)
            Set objElement2 = UndoStack(UndoStack.Count - 1)
            FrmChild.rtfText.SelStart = objElement.SelStart - (objElement.TextLen - objElement2.TextLen)
            FrmChild.rtfText.SelLength = objElement.TextLen - objElement2.TextLen
            FrmChild.rtfText.SelText = ""
            x = SendMessage(FrmChild.rtfText.hwnd, EM_HIDESELECTION, 0&, 0&)
        Else ' 增加了文本
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

' 重做

Public Sub Redo()
Dim chg As String
Dim DeleteFlag As Boolean ' 标志增加文本还是删除文本
Dim objElement As Object
    If RedoStack.Count > 0 And trapUndo Then
        trapUndo = False
        DeleteFlag = RedoStack(RedoStack.Count).TextLen < Len(FrmChild.rtfText.Text)
        If DeleteFlag Then  ' 删除了文本
            Set objElement = RedoStack(RedoStack.Count)
            FrmChild.rtfText.SelStart = objElement.SelStart
            FrmChild.rtfText.SelLength = Len(FrmChild.rtfText.Text) - objElement.TextLen
            FrmChild.rtfText.SelText = ""
        Else ' 增加了文本
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

