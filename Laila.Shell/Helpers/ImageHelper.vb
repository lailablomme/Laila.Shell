Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows
Imports System.Windows.Media.Imaging

Public Class ImageHelper
    Public Shared Function GetBitmap(source As BitmapSource) As Bitmap
        Dim bmp As Bitmap = New Bitmap(source.PixelWidth, source.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim data As BitmapData = bmp.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size),
            ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        source.CopyPixels(Int32Rect.Empty, data.Scan0, data.Height * data.Stride, data.Stride)
        bmp.UnlockBits(data)
        Return bmp
    End Function
End Class
