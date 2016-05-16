Attribute VB_Name = "FileDialog"
Option Explicit
 
Private Const OFN_READONLY = &H1  '只读方式打开
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_ALLOWMULTISELECT = &H200  '多选
Private Const OFN_EXPLORER = &H80000 '资源管理器样式（多选时会看到差别，另外貌似还可以自定样式）

Private Type OPENFILENAME
  lStructSize As Long '结构体大小
  hwndOwner As Long '调用窗体
  hInstance As Long '程序实例
  lpstrFilter As String '文件类型过滤
  lpstrCustomFilter As String
  lMaxCustFilter As Long
  lFilterIndex As Long
  lpstrFile As String '返回文件名
  lMaxFile As Long  '缓冲区大小
  lpstrFileTitle As String
  lMaxFileTitle As Long
  lpstrInitialDir As String '初始目录
  lpstrTitle As String  '对话框标题
  flags As Long '其他参数标识
  iFileOffset As Integer
  iFileExtension As Integer
  lpstrDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type

'Public Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Public Function GetFileName(hWnd As Long) As String
  Dim pFile As OPENFILENAME
  With pFile
    .lStructSize = Len(pFile)
    .hwndOwner = hWnd
    .hInstance = App.hInstance
    .lpstrFilter = "*.txt" & Chr(0) & "*.txt" & Chr(0)
    '类型说明和文件名约束用Chr(0)隔开，多种类型选择同样用Chr(0)隔开
    '同一类型多种后缀名用英文分号";"隔开，如：*.jpg;*.jpeg / *.htm;*.html
    .lMaxFile = 255
    .lpstrFile = Space(254)
    .lpstrTitle = "打开文件"
    .flags = OFN_FILEMUSTEXIST Or OFN_PATHMUSTEXIST '自己设定参数
    
  End With
  If GetOpenFileName(pFile) <> 0 Then
    GetFileName = pFile.lpstrFile
    
    GetFileName = Left(GetFileName, InStr(GetFileName, Chr(0)) - 1)
  End If
  '单文件提取：Left(函数返回值, InStr(函数返回值, Chr(0)) - 1)
  '多文件：函数返回值=所有文件所在目录(同一目录) & Chr(0) & 文件1名 & Chr(0) & 文件2名 … …
End Function

 
Public Function SaveFileName(hWnd As Long)
     
    
    Dim i As Integer
    Dim Kuang As OPENFILENAME
    Dim FileName As String
    With Kuang
        .lStructSize = Len(Kuang)
        .hwndOwner = hWnd
        .hInstance = App.hInstance
        .lpstrFile = Space(254)
        .lMaxFile = 255
        .lpstrFileTitle = Space(254)
        .lMaxFileTitle = 255
        .lpstrInitialDir = App.Path
        .flags = 6148
        '过虑对话框文件类型
        .lpstrFilter = "配置文件 (*.LevelTable)" + Chr$(0) + "*.LevelTable" + Chr$(0) '+ "配置文件 (*.LevelTable)" + Chr$(0) + "*.TXT" + Chr$(0)
        '对话框标题栏文字
        .lpstrTitle = "保存文件的路径及文件名..."
    End With
    
    i = GetSaveFileName(Kuang) '显示保存文件对话框
    If i >= 1 Then '取得对话中用户选择输入的文件名及路径
        FileName = Kuang.lpstrFile
        FileName = Left(FileName, InStr(FileName, Chr(0)) - 1)
        If (Right(FileName, 11) <> ".Yuht") Then
            FileName = FileName & ".Yuht"
        End If
    End If
    
    If Len(FileName) = 0 Then Exit Function
        SaveFileName = FileName
    Exit Function
    
End Function


