Attribute VB_Name = "mdlFile"
'**********************************************************************************'
'*     本模块有关于文件操作，包括文件新建、打开、保存及菜单中最近文件列表的更新有关文件 *'
'* 操作方面由于用了 RichTextBox 控件，所以打开文件实际上是通过新建一个窗体，然后把文件 *'
'* 导入到该控件中实现，同时该控制也限定了本程序只能支持 .txt 和 .rtf 格式式的文件。    *'
'*     最近文件列表更新的信息存储在系统注册表中，所以该相关的函数主要对注册表进行操作， *'
'* 包括读写及删除注册表项。另外由于本程序采用多文档结构，支持同时打开多个文件，而 VB 中 *'
'* 主窗口的菜单和子窗口的菜单是分开的（虽然打开子窗口时，其菜单会把主窗口的菜单替换掉） *'
'* 所以对菜单中的最近打开文档的更新要同时对两者的菜单进行更新                         *'
'**********************************************************************************'

Option Explicit

Public gMaxRecentFiles As Integer    ' 最近打开文档最多个数，其值0—9
Public gFileNameArray As Variant ' 存储文件名的数组
Private gIndex As Integer   ' 最近打开文档最大序号
Public gCountFrmChild As Integer    ' 主窗体中子窗体个数


' 检测打开的文件是否已经存在于最近打开列表中
' 若存在，返回 True ，否则，返回 False
' 函数参数为打开的文件名

Private Function OnRecentFilesList(FileName As String) As Integer
    On Error GoTo Dealerror ' 由于gFileNameArray 可能未附值而引发错误
    
    Dim i As Integer
    
    For i = 1 To UBound(gFileNameArray, 1)  ' 搜索文件列表项
        If gFileNameArray(i, 1) = FileName Then  ' 文件列表项中已存在此文件名
            gIndex = i  ' 保存该文件在文件列表中序列号，以便在其他函数中将它移到最前
            OnRecentFilesList = True
            Exit Function
        End If
    Next i
    
    OnRecentFilesList = False
    Exit Function

Dealerror:  ' 若错误，跳转到此
    OnRecentFilesList = False
End Function


' 更新文件到最近打开文档列表
' 函数参数为打开的文件名

Public Sub UpdateFileMenu(FileName As String)
    Dim intRetVal As Integer
     
    intRetVal = OnRecentFilesList(FileName)
    If Not intRetVal Then   ' 判断打开的文件名是否已经在“文件”菜单控件数组中
        WriteRecentFiles FileName, gMaxRecentFiles ' 将打开的文件名写到注册表
    ElseIf gIndex > 1 Then
        WriteRecentFiles FileName, gIndex
    End If
    GetRecentFiles  ' 更新“文件”菜单控件数组及最近打开的文件列表
End Sub

' 从注册表中读最近打开文档列表记录并更新到菜单中

Public Sub GetRecentFiles()
    Dim i As Integer    ' 记数
    
    If GetSetting(App.Title, "Recent Files", "Recent File1") = Empty Then
        ' 检测注册表中相应项目是否为空
        FrmMain.mnuFileRec(0).Visible = True    ' 显示菜单中空的最近打开文档列表
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileRec(0).Visible = True
        End If
        For i = 1 To gMaxRecentFiles
            FrmMain.mnuFileRec(i).Visible = False   ' 隐藏已经显示的最近打开文档
            If gCountFrmChild > 0 Then
                FrmMain.ActiveForm.mnuFileRec(i).Visible = False
            End If
        Next i
        FrmMain.mnuFileSeperator3.Visible = False  ' 隐藏文件列表项与退出菜单之间的分隔条
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileSeperator3.Visible = False
        End If
        
        gFileNameArray = Empty
        Exit Sub
    End If
    
    gFileNameArray = GetAllSettings(App.Title, "Recent Files")  ' 从注册表中得到最近打开文档信息
    For i = 1 To UBound(gFileNameArray, 1)   ' 获取数组最大下标
        ' 获得文件名及路径
        FrmMain.mnuFileRec(i).Caption = "&" & Format$(i) & " " & gFileNameArray(i, 1)
        FrmMain.mnuFileRec(i).Visible = True   ' 显示注册表中存在的最近打开文档
        If gCountFrmChild > 0 Then
            FrmMain.ActiveForm.mnuFileRec(i).Caption = "&" & Format$(i) & " " & gFileNameArray(i, 1)
            FrmMain.ActiveForm.mnuFileRec(i).Visible = True
        End If
    Next i
    FrmMain.mnuFileRec(0).Visible = False
    FrmMain.mnuFileSeperator3.Visible = True  ' 显示文件列表项与退出菜单之间的分隔条
    If gCountFrmChild > 0 Then
        FrmMain.ActiveForm.mnuFileRec(0).Visible = False
        FrmMain.ActiveForm.mnuFileSeperator3.Visible = True
    End If
End Sub
  
' 把最近打开文档信息写入注册表
' 第一个参数为文件名，第二个为其在最近打开文档中的序列号

Private Sub WriteRecentFiles(FileName As String, index As Integer)
    Dim i As Integer    ' 计数
    Dim strFileName As String   ' 文件名
    Dim strKey As String    ' 在注册表中的关键字项
    
    ' 将文件Recent File1 复制给Recent File2...，使文件列表项顺次下移
    For i = index - 1 To 1 Step -1
        strKey = "Recent File" & i
        ' 从Windows 注册表中的应用程序项目返回注册表项设置值
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            strKey = "Recent File" & (i + 1)
            SaveSetting App.Title, "Recent Files", strKey, strFileName '写入注册表
        End If
    Next i
    ' 将最新打开的文件写到最近使用文件列表的第一项
    SaveSetting App.Title, "Recent Files", "Recent File1", FileName
End Sub

' 把指定最近打开文档信息从注册表删除
' 参数为其在最近打开文档中的序列号

Public Sub DeleteRecentFile(index As Integer)
    Dim i As Integer    ' 计数
    Dim strFileName As String   ' 文件名
    Dim strKey As String    ' 在注册表中的关键字项
    
    If index > gMaxRecentFiles Then ' 如果不可能在注册表中存在
        Exit Sub
    End If
    
    ' 将文件列表项顺次上移到第 index 项，最后删除最后一项
    For i = index + 1 To gMaxRecentFiles
        strKey = "Recent File" & i
        ' 从Windows 注册表中的应用程序项目返回注册表项设置值
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            strKey = "Recent File" & (i - 1)
            SaveSetting App.Title, "Recent Files", strKey, strFileName '写入注册表
        Else
            Exit For
        End If
    Next i
    
    strKey = "recent file" & (i - 1)
    DeleteSetting App.Title, "Recent Files", strKey ' 删除最后一项
End Sub

' 把所有最近打开文档信息从注册表删除

Public Sub DeleteAllRecentFiles()
    Dim i As Integer    ' 计数
    Dim strFileName As String   ' 文件名
    Dim strKey As String    ' 在注册表中的关键字项

    ' 逐项删除
    For i = 1 To gMaxRecentFiles
        strKey = "Recent File" & i
        ' 从Windows 注册表中的应用程序项目返回注册表项设置值
        strFileName = GetSetting(App.Title, "Recent Files", strKey)
        If strFileName <> "" Then
            DeleteSetting App.Title, "Recent Files", strKey ' 删除注册表中信息
        Else
            Exit For
        End If
    Next i
End Sub

' 创建新文件

Public Sub NewFile()
    Dim frmNew As New FrmChild  ' 创建新窗体对象
    Static lgFileCount As Long  ' 记录新建的文件个数
    
    lgFileCount = lgFileCount + 1   ' 计数增加
    frmNew.Caption = "Noname" & lgFileCount ' 将新窗体的标题设为“Noname”+ 数字
    frmNew.Show ' 显示窗体
    frmNew.m_Modified = False ' 设置标志为未修改
End Sub

' 打开新文件，通过先创建新文件窗体，再把文件读到窗体中的 RichTextBox 中实现

Public Sub OpenFile()
    On Error Resume Next    ' 设置可以跳过错误
    
    Dim strFileName As String   ' 文件名
    
    With FrmMain.dlgCommonDialog    ' 打开文件通用对话框的初始设置
        .DialogTitle = "Open"
        .Filter = "Text (*.txt)|*.txt|RTFText (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowOpen
        If Err = 32755 Then ' 用户点击 Cancel 后退出
            Exit Sub
        End If
        If Len(.FileName) = 0 Then  ' 检验文件名
            Exit Sub
        End If
        strFileName = .FileName
    End With
        
    Screen.MousePointer = 11    ' 设置鼠标为等待形状
    
    NewFile  ' 新建新文档窗体
    FrmMain.ActiveForm.rtfText.LoadFile strFileName ' RichTextBox 中加载文件
    FrmMain.ActiveForm.Caption = strFileName    ' 窗体标题为该文件名
    FrmMain.ActiveForm.m_Modified = False ' 设置标志为未修改
    
    Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
    
    UpdateFileMenu (strFileName)    ' 更新最近打开文档列表
End Sub

' 保存文件

Public Sub SaveFile()
    On Error Resume Next    ' 设置可以跳过错误
    
    Dim strFileName As String   '取得保存文件名，保存文件
    
    If Left(FrmMain.ActiveForm.Caption, 6) = "Noname" Then  ' 若文件未保存过，则调用另存为
        SaveFileAs
    Else
        Screen.MousePointer = 11    ' 设置鼠标为等待形状
        
        strFileName = FrmMain.ActiveForm.Caption    ' 窗体标题为该文件名
        FrmMain.ActiveForm.rtfText.SaveFile strFileName ' 保存文件
        FrmMain.ActiveForm.m_Modified = False ' 设置标志为未修改
        
        Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
    End If
End Sub

' 文件另存为

Public Sub SaveFileAs()
    On Error Resume Next    ' 设置可以跳过错误
    
    Dim strFileName As String   ' 文件名
    
    With FrmMain.dlgCommonDialog    ' 另存为通用对话框架的实现
        .DialogTitle = "Save As"
        .CancelError = False
        .Filter = "Text (*.txt)|*.txt|RTFText (*.rtf)|*.rtf|All Files (*.*)|*.*"
        .FilterIndex = 1
        .ShowSave
        If Len(.FileName) = 0 Then  ' 检验文件名
            Exit Sub
        End If
        strFileName = .FileName
    End With
    
    Screen.MousePointer = 11    ' 设置鼠标为等待形状
    
    FrmMain.ActiveForm.Caption = strFileName    ' 窗体标题为该文件名
    FrmMain.ActiveForm.rtfText.SaveFile strFileName ' 保存文件
    FrmMain.ActiveForm.m_Modified = False ' 设置标志为未修改
    
    Screen.MousePointer = 0 ' 恢复鼠标为通常的箭头形状
    
    UpdateFileMenu (strFileName)    ' 更新最近打开文档列表
End Sub

