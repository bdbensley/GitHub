Imports System.Drawing
Imports System.Drawing.Drawing2D

Public Class Form1

    Public Shared Image As Image
    Public Shared tempImage As Image
    Public Shared Bound As Rectangle
    Dim button2clicked As Boolean = False
    Public Shared ScreenNumber As Integer = 0
    Public Shared W As Integer = Screen.AllScreens(ScreenNumber).Bounds.Width
    Public Shared H As Integer = Screen.AllScreens(ScreenNumber).Bounds.Height
    Dim screen1 As Screen
    Dim screen0 As Screen

    Shared Function GetDesktopImage(Optional ByVal Width As Integer = 0, Optional ByVal Height As Integer = 0, Optional ByVal ShowCursor As Boolean = True) As Image

        Dim DesktopBitmap As New Bitmap(W, H) ' width, height
        Dim g As Graphics = Graphics.FromImage(DesktopBitmap)
        Dim screen0 As Screen
        screen0 = Screen.AllScreens(0)
        g.CopyFromScreen(100, 100, 0, 0, New Size(1000, 800))
        'If ShowCursor Then Cursors.Default.Draw(g, New Rectangle(New Point(Cursor.Position.X - Screen.AllScreens(0).Bounds.Width, Cursor.Position.Y), New Size(32, 32)))
        g.Dispose()
        'Width = 0
        'Height = 0
        Image = DesktopBitmap.Clone
        'If Width = 0 And Height = 0 Then
        'Return Image 'DesktopBitmap
        'Else
        'Dim ScaledBitmap As Image = DesktopBitmap.GetThumbnailImage(Width, Height, Nothing, IntPtr.Zero)
        'Image = DesktopBitmap.Clone
        'Image = ScaledBitmap.Clone
        'ResizeImage(ScaledBitmap.Clone, New Size(1920, 1080), False)
        'Dim newWidth As Integer
        'Dim newHeight As Integer
        'Dim originalWidth As Integer = 800
        'Dim originalHeight As Integer = 800
        'Dim percentWidth As Single = CSng(Width) / CSng(originalWidth)
        'Dim percentHeight As Single = CSng(Height) / CSng(originalHeight)
        'Dim percent As Single = If(percentHeight < percentWidth, percentHeight, percentWidth)
        'newWidth = CInt(originalWidth * percentWidth)
        'newHeight = CInt(originalHeight * percentHeight)


        'Dim newImage As Image = New Bitmap(newWidth, newHeight)


        'Using graphicsHandle As Graphics = Graphics.FromImage(newImage)
        'graphicsHandle.InterpolationMode = InterpolationMode.HighQualityBicubic
        'graphicsHandle.DrawImage(Image, 0, 0, newWidth, newHeight)
        'End Using
        Return Image 'ScaledBitmap
        'End If

    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        screen1 = Screen.AllScreens(1)
        Me.StartPosition = FormStartPosition.Manual
        Me.Location = screen1.Bounds.Location + New Point(0, 0)
        Me.WindowState = FormWindowState.Maximized
        Me.Show()
        screen0 = Screen.AllScreens(0)
        Application.DoEvents()
        Do Until button2clicked = True
            Application.DoEvents()
            GC.Collect()
            tempImage = GetDesktopImage()
            PictureBox1.Image = tempImage
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
        Loop
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        button2clicked = True
        End
    End Sub
End Class
