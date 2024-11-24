Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports FileSignatures
Imports FileSignatures.Formats
Imports Laila.Shell.Helpers

Public Class ImageHelper
    Private Shared _recognised As IEnumerable(Of FileSignatures.Formats.Image)
    Private Shared _inspector As FileFormatInspector
    Private Shared _icons As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()
    Private Shared _icons2 As Dictionary(Of Integer, BitmapSource) = New Dictionary(Of Integer, BitmapSource)()
    Private Shared _overlayIconIndexes As Dictionary(Of Byte, Integer) = New Dictionary(Of Byte, Integer)()
    Private Shared _imageList As IImageList


    Shared Sub New()
        _recognised = FileFormatLocator.GetFormats().OfType(Of FileSignatures.Formats.Image)()
        _inspector = New FileFormatInspector(_recognised)

        UIHelper.OnUIThread(
            Sub()
                Dim ptr As IntPtr
                Try
                    Functions.SHGetImageList(SHIL.SHIL_EXTRALARGE, GetType(IImageList).GUID, ptr)
                    _imageList = Marshal.GetObjectForIUnknown(ptr)
                Finally
                    If Not IntPtr.Zero.Equals(ptr) Then
                        Marshal.Release(ptr)
                    End If
                End Try
            End Sub)
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

    Public Shared Function GetOverlayIcon(overlayIconIndex As Byte) As BitmapSource
        If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
            Dim index As Integer
            UIHelper.OnUIThread(
                Sub()
                    If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
                        _imageList.GetOverlayImage(overlayIconIndex, index)
                        _overlayIconIndexes.Add(overlayIconIndex, index)
                    End If
                End Sub, Threading.DispatcherPriority.Send)
        End If

        Return ImageHelper.GetIcon(_overlayIconIndexes(overlayIconIndex))
    End Function

    Public Shared Function GetIcon(index As Integer) As BitmapSource
        If Not _icons2.ContainsKey(index) Then
            UIHelper.OnUIThread(
                Sub()
                    If Not _icons2.ContainsKey(index) Then
                        Dim hIcon As IntPtr
                        Try
                            _imageList.GetIcon(index, 0, hIcon)
                            Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                                Dim image As BitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(icon.ToBitmap().GetHbitmap(), IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                image.Freeze()
                                _icons2.Add(index, image)
                            End Using
                        Finally
                            If Not IntPtr.Zero.Equals(hIcon) Then
                                Functions.DestroyIcon(hIcon)
                            End If
                        End Try
                    End If
                End Sub, Threading.DispatcherPriority.Send)
        End If

        Return _icons2(index)
    End Function

    Public Shared Function ExtractIcon(ref As String) As BitmapSource
        If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
            Dim s() As String = Split(ref, ","), icon As IntPtr, iconl As IntPtr
            Try
                Functions.ExtractIconEx(s(0), s(1), iconl, icon, 1)
                If Not IntPtr.Zero.Equals(icon) Then
                    Dim img As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    img.Freeze()
                    UIHelper.OnUIThread(
                        Sub()
                            If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
                                _icons.Add(ref.ToLower().Trim(), img)
                            End If
                        End Sub, Threading.DispatcherPriority.Send)
                Else
                    UIHelper.OnUIThread(
                        Sub()
                            If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
                                _icons.Add(ref.ToLower().Trim(), Nothing)
                            End If
                        End Sub, Threading.DispatcherPriority.Send)
                End If
            Finally
                If Not IntPtr.Zero.Equals(icon) Then
                    Functions.DestroyIcon(icon)
                End If
                If Not IntPtr.Zero.Equals(iconl) Then
                    Functions.DestroyIcon(iconl)
                End If
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
