Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports FileSignatures
Imports FileSignatures.Formats

Public Class ImageHelper
    Private Shared _recognised As IEnumerable(Of FileSignatures.Formats.Image)
    Private Shared _inspector As FileFormatInspector
    Private Shared _icons As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()

    Shared Sub New()
        _recognised = FileFormatLocator.GetFormats().OfType(Of FileSignatures.Formats.Image)()
        _inspector = New FileFormatInspector(_recognised)
    End Sub

    Public Shared Function IsImage(fullPath As String) As Boolean
        Try
            Using fs = New FileStream(fullPath, FileMode.Open, FileAccess.Read)
                Return TypeOf _inspector.DetermineFileFormat(fs) Is Formats.Image
            End Using
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function ExtractIcon(ref As String) As BitmapSource
        If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
            Dim s() As String = Split(ref, ","), icon As IntPtr, iconl As IntPtr
            Try
                Functions.ExtractIconEx(s(0), s(1), iconl, Icon, 1)
                If Not IntPtr.Zero.Equals(Icon) Then
                    _icons.Add(ref.ToLower().Trim(), System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions()))
                Else
                    _icons.Add(ref.ToLower().Trim(), Nothing)
                End If
            Finally
                Functions.DeleteObject(Icon)
                Functions.DeleteObject(iconl)
            End Try
        End If
        Return _icons(ref.ToLower().Trim())
    End Function

    Public Shared Function GetBitmap(source As BitmapSource) As Bitmap
        Dim bmp As Bitmap = New Bitmap(source.PixelWidth, source.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim data As BitmapData = bmp.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size),
            ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        source.CopyPixels(Int32Rect.Empty, data.Scan0, data.Height * data.Stride, data.Stride)
        bmp.UnlockBits(data)
        Return bmp
    End Function
End Class
