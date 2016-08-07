Imports System.IO
Imports Microsoft.VisualBasic.FileIO

Public Class RecycleBinForm
    Private Declare Function SetParent Lib "user32" Alias "SetParent" (ByVal hWndChild As Integer, ByVal hWndNewParent As Integer) As Integer
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Integer, ByVal hWnd2 As Integer, ByVal lpsz1 As String, ByVal lpsz2 As String) As Integer
    Private Declare Function GetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Integer
    Private Declare Function ReleaseCapture Lib "user32" () As Integer
    Private Declare Function SendMessageA Lib "user32" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As VariantType) As Integer

    ''' <summary>
    ''' 清空回收站
    ''' </summary>
    ''' <param name="hwnd">父级句柄</param>
    ''' <param name="pszRootPath">为空则清空所有驱动器的回收站</param>
    ''' <param name="dwFlags">参数标识</param>
    Private Declare Sub SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As IntPtr, ByVal pszRootPath As String, ByVal dwFlags As Integer)
    Private Const SHERB_NOCONFIRMATION = &H1 '不显示确认提示框
    Private Const SHERB_NOPROGRESSUI = &H2 '不显示任务进度条
    Private Const SHERB_NOSOUND = &H4 '不播放清空提示音
    Public FileIcon As Bitmap, FileName As String

    ''' <summary>
    ''' 更新回收站图标
    ''' </summary>
    Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Integer
    'API操作成功返回值
    Const S_OK = &H0

    '黑洞每次旋转的度数
    Private Const dAngle As Integer = 20
    '黑洞下一次的旋转角度
    Dim Angle As Integer = dAngle

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        BitmapModule.SetAlphaPicture(Me, BitmapModule.GetRotateBitmap(My.Resources.RecycleBinRes.BlackHole, Angle))
        Angle = IIf(Angle = 360, dAngle, Angle + dAngle)
    End Sub
    Protected Overloads Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            If Not DesignMode Then
                Dim cp As CreateParams = MyBase.CreateParams
                cp.ExStyle = cp.ExStyle Or BitmapModule.WS_EX_LAYERED
                Return cp
            Else
                Return MyBase.CreateParams
            End If
        End Get
    End Property

    Private Sub RecycleBinForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '允许拖拽
        Me.AllowDrop = True
        SetParent(Me.Handle, GetDesktopIconHandle)
        BitmapModule.SetAlphaPicture(Me, My.Resources.RecycleBinRes.BlackHole)
        Me.Location = New Point(My.Computer.Screen.Bounds.Width - Me.Width, 10)
    End Sub

    Private Sub RecycleBinForm_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim FilePaths() As String = e.Data.GetData(DataFormats.FileDrop)
        For Each File As String In FilePaths
            Try
                FileIcon = GetFileIcon(File)
                FileName = FileSystem.GetName(File)
                If FileSystem.FileExists(File) Then
                    FileSystem.DeleteFile(File, UIOption.AllDialogs, RecycleOption.SendToRecycleBin)
                ElseIf FileSystem.DirectoryExists(File) Then
                    FileSystem.DeleteDirectory(File, UIOption.AllDialogs, RecycleOption.SendToRecycleBin)
                End If
            Finally
                '执行完毕
            End Try
        Next
        FileIcon = Nothing
        '停止旋转
        Timer1.Stop()
    End Sub

    Public Function GetDesktopIconHandle() As IntPtr
        '获取物理系统桌面图标的句柄，用于嵌入实现置后显示
        Dim HandleDesktop As Integer = GetDesktopWindow
        Dim HandleTop As Integer = 0
        Dim LastHandleTop As Integer = 0
        Dim HandleSHELLDLL_DefView As Integer = 0
        Dim HandleSysListView32 As Integer = 0
        '在WorkerW结构里搜索
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "WorkerW", vbNullString)
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do
        Loop
        '如果找到了，立即返回
        If HandleSysListView32 > 0 Then Return HandleSysListView32
        '未找到，则在Progman里搜索(用于兼容WinXP系统)
        Do Until HandleSysListView32 > 0
            HandleTop = FindWindowEx(HandleDesktop, LastHandleTop, "Progman", "Program Manager")
            HandleSHELLDLL_DefView = FindWindowEx(HandleTop, 0, "SHELLDLL_DefView", vbNullString)
            If HandleSHELLDLL_DefView > 0 Then HandleSysListView32 = FindWindowEx(HandleSHELLDLL_DefView, 0, "SysListView32", "FolderView")
            LastHandleTop = HandleTop
            If LastHandleTop = 0 Then Exit Do : Return 0
        Loop
        Return HandleSysListView32
    End Function

    Private Sub RecycleBinForm_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If (e.Data.GetDataPresent(DataFormats.FileDrop) = True) Then e.Effect = DragDropEffects.All : Timer1.Start()
    End Sub

    Private Sub RecycleBinForm_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        If e.Button = MouseButtons.Right Then End
    End Sub

    Private Sub RecycleBinForm_DragLeave(sender As Object, e As EventArgs) Handles Me.DragLeave
        Timer1.Stop()
    End Sub

    Private Sub RecycleBinForm_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        ReleaseCapture()
        SendMessageA(Me.Handle, &HA1, 2, 0&)
    End Sub
End Class