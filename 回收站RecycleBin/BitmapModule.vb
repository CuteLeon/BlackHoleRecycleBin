Imports System.Drawing.Drawing2D
Imports System.Runtime.InteropServices

Module BitmapModule
    Private Const ULW_COLORKEY As Int32 = &H1
    Public Const ULW_ALPHA As Int32 = &H2
    Private Const ULW_OPAQUE As Int32 = &H4
    Public Const WS_EX_LAYERED As Int32 = &H80000
    Public Const AC_SRC_OVER As Byte = &H0
    Public Const AC_SRC_ALPHA As Byte = &H1

    <StructLayout(LayoutKind.Sequential)>
    Public Structure mSize
        Private cx As Int32
        Private cy As Int32
        Public Sub New(ByVal cx As Int32, ByVal cy As Int32)
            Me.cx = cx
            Me.cy = cy
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure mPoint
        Private x As Int32
        Private y As Int32
        Public Sub New(ByVal x As Int32, ByVal y As Int32)
            Me.x = x
            Me.y = y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Sequential, Pack:=1)>
    Public Structure BLENDFUNCTION
        Public BlendOp As Byte
        Public BlendFlags As Byte
        Public SourceConstantAlpha As Byte
        Public AlphaFormat As Byte
    End Structure

    Public Declare Auto Function GetDC Lib "user32.dll" (ByVal hWnd As IntPtr) As IntPtr
    Public Declare Auto Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As IntPtr) As IntPtr
    Public Declare Auto Function SelectObject Lib "gdi32.dll" (ByVal hDC As IntPtr, ByVal hObject As IntPtr) As IntPtr
    Public Declare Auto Function UpdateLayeredWindow Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal hdcDst As IntPtr, ByRef pptDst As mPoint, ByRef psize As mSize, ByVal hdcSrc As IntPtr, ByRef pprSrc As mPoint, ByVal crKey As Int32, ByRef pblend As BLENDFUNCTION, ByVal dwFlags As Int32) As Boolean
    Public Declare Auto Function ReleaseDC Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal hDC As IntPtr) As Integer
    Public Declare Auto Function DeleteObject Lib "gdi32.dll" (ByVal hObject As IntPtr) As Boolean
    Public Declare Auto Function DeleteDC Lib "gdi32.dll" (ByVal hdc As IntPtr) As Boolean

    Dim hDC1 As IntPtr = GetDC(IntPtr.Zero)
    Dim hDC2 As IntPtr = CreateCompatibleDC(hDC1)
    Dim hBitmap1 As IntPtr = IntPtr.Zero
    Dim hBitmap2 As IntPtr = IntPtr.Zero
    Dim blend As New BLENDFUNCTION()

    Dim GetBitmap As Bitmap
    Dim ReturnBitmap As Bitmap
    Dim MyGraphics As Graphics

    ' 使用Alpha通道把Bitmap绘制到Window
    Public Sub SetAlphaPicture(ByVal AlphaWindow As Form, ByVal AlphaPicture As Bitmap)
        Try
            hBitmap1 = AlphaPicture.GetHbitmap(Color.FromArgb(0))
            hBitmap2 = SelectObject(hDC2, hBitmap1)

            With blend
                .BlendOp = AC_SRC_OVER
                .BlendFlags = 0
                .AlphaFormat = AC_SRC_ALPHA
                .SourceConstantAlpha = True
            End With
            Call UpdateLayeredWindow(AlphaWindow.Handle, hDC1, New mPoint(AlphaWindow.Left, AlphaWindow.Top), New mSize(AlphaPicture.Width, AlphaPicture.Height), hDC2, New mPoint(0, 0), 0, blend, ULW_ALPHA)
        Finally
            Call SelectObject(hDC2, hBitmap2)
            Call DeleteObject(hBitmap1)
        End Try
    End Sub

    '旋转Bitmap图像
    Public Function GetRotateBitmap(ByVal BitmapRes As Bitmap, ByVal Angle As Single) As Bitmap
        GetBitmap = BitmapRes
        If Not ReturnBitmap Is Nothing Then ReturnBitmap.Dispose()
        ReturnBitmap = New Bitmap(GetBitmap.Width, GetBitmap.Height)
        MyGraphics = Graphics.FromImage(ReturnBitmap)
        MyGraphics.TranslateTransform(GetBitmap.Width / 2, GetBitmap.Height / 2)
        MyGraphics.RotateTransform(Angle, MatrixOrder.Prepend)
        MyGraphics.TranslateTransform(-GetBitmap.Width / 2, -GetBitmap.Height / 2)
        MyGraphics.DrawImage(GetBitmap, 0, 0, GetBitmap.Width, GetBitmap.Height)
        If Not RecycleBinForm.FileIcon Is Nothing Then
            MyGraphics.DrawImage(RecycleBinForm.FileIcon, 112, 112, 32, 32)
            MyGraphics.DrawString(RecycleBinForm.FileName, RecycleBinForm.Font, Brushes.White, 112, 150)
        End If
        MyGraphics.Dispose()
        Return ReturnBitmap
    End Function

    Private Structure SHFILEINFO
        Public hIcon As IntPtr
        Public iIcon As Integer
        Public dwAttributes As Integer
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public szDisplayName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=80)>
        Public szTypeName As String
    End Structure

    Private Declare Auto Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr
    Private Const SHGFI_ICON = &H100
    Private Const SHGFI_LARGEICON = &H0

    Public Function GetFileIcon(ByVal FilePath As String) As Bitmap
        Dim SHInfo As New SHFILEINFO()
        SHInfo.szDisplayName = New String(Chr(0), 260)
        SHInfo.szTypeName = New String(Chr(0), 80)
        SHGetFileInfo(FilePath, 0, SHInfo, Marshal.SizeOf(SHInfo), SHGFI_ICON Or SHGFI_LARGEICON)
        Dim FileIcon As Bitmap = System.Drawing.Icon.FromHandle(SHInfo.hIcon).ToBitmap
        Return FileIcon
    End Function
End Module
