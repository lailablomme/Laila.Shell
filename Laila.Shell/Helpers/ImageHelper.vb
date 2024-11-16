Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows
Imports System.Windows.Media.Imaging
Imports FileSignatures

Public Class ImageHelper
    Private Shared _recognised As IEnumerable(Of FileSignatures.Formats.Image)
    Private Shared _inspector As FileFormatInspector

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

    Public Shared Function GetBitmap(source As BitmapSource) As Bitmap
        Dim bmp As Bitmap = New Bitmap(source.PixelWidth, source.PixelHeight, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        Dim data As BitmapData = bmp.LockBits(New System.Drawing.Rectangle(System.Drawing.Point.Empty, bmp.Size),
            ImageLockMode.WriteOnly, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
        source.CopyPixels(Int32Rect.Empty, data.Scan0, data.Height * data.Stride, data.Stride)
        bmp.UnlockBits(data)
        Return bmp
    End Function
End Class
