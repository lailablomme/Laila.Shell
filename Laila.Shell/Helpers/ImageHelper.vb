Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Forms
Imports System.Windows.Media.Imaging
Imports FileSignatures
Imports FileSignatures.Formats
Imports Laila.Shell.Helpers

Public Class ImageHelper
    Private Shared _recognised As IEnumerable(Of FileSignatures.Formats.Image)
    Private Shared _inspector As FileFormatInspector
    Private Shared _iconsLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
    Private Shared _icons As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()
    Private Shared _icons2 As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()
    Private Shared _overlayIconIndexes As Dictionary(Of Byte, Integer) = New Dictionary(Of Byte, Integer)()
    Private Shared _imageListSmall As IImageList
    Private Shared _imageListLarge As IImageList
    Private Shared _imageListExtraLarge As IImageList
    Private Shared _imageListJumbo As IImageList


    Shared Sub New()
        _recognised = FileFormatLocator.GetFormats().OfType(Of FileSignatures.Formats.Image)()
        _inspector = New FileFormatInspector(_recognised)
    End Sub

    Public Shared Sub Load()
        Functions.SHGetImageList(SHIL.SHIL_SMALL, GetType(IImageList).GUID, _imageListSmall)
        Functions.SHGetImageList(SHIL.SHIL_LARGE, GetType(IImageList).GUID, _imageListLarge)
        Functions.SHGetImageList(SHIL.SHIL_EXTRALARGE, GetType(IImageList).GUID, _imageListExtraLarge)
        Functions.SHGetImageList(SHIL.SHIL_JUMBO, GetType(IImageList).GUID, _imageListJumbo)
    End Sub

    Public Shared Function IsImage(fullPath As String) As Boolean
        If IO.File.Exists(fullPath) Then
            Try
                Using fs = New FileStream(fullPath, FileMode.Open, FileAccess.Read)
                    Return TypeOf _inspector.DetermineFileFormat(fs) Is Formats.Image
                End Using
            Catch ex As Exception
            End Try
        End If
        Return False
    End Function

    Public Shared Function GetOverlayIcon(overlayIconIndex As Byte, size As Integer) As BitmapSource
        If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
            Dim index As Integer
            UIHelper.OnUIThread(
                Sub()
                    If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
                        _imageListSmall.GetOverlayImage(overlayIconIndex, index)
                        _overlayIconIndexes.Add(overlayIconIndex, index)
                    End If
                End Sub, Threading.DispatcherPriority.Send)
        End If

        Return ImageHelper.GetIcon(_overlayIconIndexes(overlayIconIndex), size)
    End Function

    Public Shared Function GetIcon(index As Integer, size As Integer) As BitmapSource
        If Not _icons2.ContainsKey(String.Format("{0}_{1}", index, size)) Then
            UIHelper.OnUIThread(
                Sub()
                    If Not _icons2.ContainsKey(String.Format("{0}_{1}", index, size)) Then
                        Dim hIcon As IntPtr
                        Try
                            If size <= 16 Then
                                _imageListSmall.GetIcon(index, 0, hIcon)
                            ElseIf size <= 32 Then
                                _imageListLarge.GetIcon(index, 0, hIcon)
                            ElseIf size <= 48 Then
                                _imageListExtraLarge.GetIcon(index, 0, hIcon)
                            Else
                                _imageListJumbo.GetIcon(index, 0, hIcon)
                            End If
                            Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                                Using bitmap = icon.ToBitmap()
                                    Dim hBitmap As IntPtr = bitmap.GetHbitmap()
                                    Dim image As BitmapSource = Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                    image.Freeze()
                                    _icons2.Add(String.Format("{0}_{1}", index, size), image)
                                End Using
                            End Using
                        Finally
                            If Not IntPtr.Zero.Equals(hIcon) Then
                                Functions.DestroyIcon(hIcon)
                            End If
                        End Try
                    End If
                End Sub, Threading.DispatcherPriority.Send)
        End If

        Return _icons2(String.Format("{0}_{1}", index, size))
    End Function

    Public Shared Function ExtractIcon(ref As String) As BitmapSource
        If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
            Dim s() As String = Split(ref, ","), icon As IntPtr, iconl As IntPtr
            Try
                Functions.ExtractIconEx(s(0), s(1), iconl, icon, 1)
                If Not IntPtr.Zero.Equals(icon) Then
                    Dim img As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(icon, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                    img.Freeze()
                    _iconsLock.Wait()
                    Try
                        If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
                            _icons.Add(ref.ToLower().Trim(), img)
                        End If
                    Finally
                        _iconsLock.Release()
                    End Try
                Else
                    _iconsLock.Wait()
                    Try
                        If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
                            _icons.Add(ref.ToLower().Trim(), Nothing)
                        End If
                    Finally
                        _iconsLock.Release()
                    End Try
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
