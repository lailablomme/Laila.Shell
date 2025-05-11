Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Globalization
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Items

Namespace Helpers
    Public Class ImageHelper
        Private Shared _iconsLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
        Private Shared _icons As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()
        Private Shared _icons2 As Dictionary(Of String, BitmapSource) = New Dictionary(Of String, BitmapSource)()
        Private Shared _overlayIconIndexes As Dictionary(Of Byte, Integer) = New Dictionary(Of Byte, Integer)()
        Private Shared _overlayIconLock As Object = New Object()
        Private Shared _imageListSmall As IImageList
        Private Shared _imageListLarge As IImageList
        Private Shared _imageListExtraLarge As IImageList
        Private Shared _imageListJumbo As IImageList

        Public Shared Property DefaultFileIconSmall As ImageSource
        Public Shared Property DefaultFileIconLarge As ImageSource
        Public Shared Property DefaultFileIconExtraLarge As ImageSource
        Public Shared Property DefaultFileIconJumbo As ImageSource
        Public Shared Property DefaultFolderIconSmall As ImageSource
        Public Shared Property DefaultFolderIconLarge As ImageSource
        Public Shared Property DefaultFolderIconExtraLarge As ImageSource
        Public Shared Property DefaultFolderIconJumbo As ImageSource

        Private Shared _imageFileExtensions As String() = {
            ".bmp", ".dib",    ' Bitmap formats
            ".jpg", ".jpeg", ".jpe", ".jfif", ".jfi",    ' JPEG formats
            ".png",            ' Portable Network Graphics
            ".gif",            ' Graphics Interchange Format
            ".tif", ".tiff",   ' Tagged Image File Format
            ".ico",            ' Icon files
            ".webp",           ' WebP format (Windows 10+)
            ".heif", ".heic"   ' High-Efficiency Image Format (Windows 10+ with extensions)
        }


        Shared Sub New()
        End Sub

        Private Shared Function getFileIcon(attributes As Integer, size As Integer) As ImageSource
            Dim shfi As New SHFILEINFO(), result As BitmapSource
            Try
                Functions.SHGetFileInfo(
                    "c:\dummy",
                    attributes,
                    shfi,
                    Marshal.SizeOf(shfi),
                    SHGFI.SHGFI_ICON Or SHGFI.SHGFI_USEFILEATTRIBUTES)
                Return ImageHelper.GetIcon(shfi.iIcon, size)
            Finally
                If Not IntPtr.Zero.Equals(shfi.hIcon) Then
                    Functions.DestroyIcon(shfi.hIcon)
                    shfi.hIcon = IntPtr.Zero
                End If
            End Try
            Return result
        End Function

        Public Shared Sub Load()
            Functions.SHGetImageList(SHIL.SHIL_SMALL, GetType(IImageList).GUID, _imageListSmall)
            Functions.SHGetImageList(SHIL.SHIL_LARGE, GetType(IImageList).GUID, _imageListLarge)
            Functions.SHGetImageList(SHIL.SHIL_EXTRALARGE, GetType(IImageList).GUID, _imageListExtraLarge)
            Functions.SHGetImageList(SHIL.SHIL_JUMBO, GetType(IImageList).GUID, _imageListJumbo)

            DefaultFileIconSmall = getFileIcon(&H80, 16)
            DefaultFileIconLarge = getFileIcon(&H80, 32)
            DefaultFileIconExtraLarge = getFileIcon(&H80, 48)
            DefaultFileIconJumbo = getFileIcon(&H80, 128)
            DefaultFolderIconSmall = getFileIcon(&H10, 16)
            DefaultFolderIconLarge = getFileIcon(&H10, 32)
            DefaultFolderIconExtraLarge = getFileIcon(&H10, 48)
            DefaultFolderIconJumbo = getFileIcon(&H10, 128)
        End Sub

        Public Shared Function GetApplicationIcon(path As String) As ImageSource
            Dim hBitmap As IntPtr
            Try
                Using icon As System.Drawing.Icon = Icon.ExtractAssociatedIcon(path)
                    Using bitmap = icon.ToBitmap()
                        hBitmap = bitmap.GetHbitmap()
                        Dim image As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                        image.Freeze()
                        Return image
                    End Using
                End Using
            Catch ex As Exception
                ' ignore
            Finally
                If Not IntPtr.Zero.Equals(hBitmap) Then
                    Functions.DeleteObject(hBitmap)
                    hBitmap = IntPtr.Zero
                End If
            End Try
        End Function

        Public Shared Function IsImage(fullPath As String) As Boolean
            Return _imageFileExtensions.Contains(IO.Path.GetExtension(fullPath)?.ToLower())
        End Function

        Public Shared Function GetOverlayIcon(overlayIconIndex As Byte, size As Integer) As BitmapSource
            If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
                Dim index As Integer
                UIHelper.OnUIThread(
                    Sub()
                        SyncLock _overlayIconLock
                            If Not _overlayIconIndexes.ContainsKey(overlayIconIndex) Then
                                _imageListSmall.GetOverlayImage(overlayIconIndex, index)
                                _overlayIconIndexes.Add(overlayIconIndex, index)
                            End If
                        End SyncLock
                    End Sub)
            End If

            Return ImageHelper.GetIcon(_overlayIconIndexes(overlayIconIndex), size)
        End Function

        Public Shared Function GetIcon(index As Integer, size As Integer) As BitmapSource
            If Not _icons2.ContainsKey(String.Format("{0}_{1}", index, size)) Then
                Dim hIcon As IntPtr
                UIHelper.OnUIThread(
                    Sub()
                        If size <= 16 Then
                            _imageListSmall.GetIcon(index, 0, hIcon)
                        ElseIf size <= 32 Then
                            _imageListLarge.GetIcon(index, 0, hIcon)
                        ElseIf size <= 48 Then
                            _imageListExtraLarge.GetIcon(index, 0, hIcon)
                        Else
                            _imageListJumbo.GetIcon(index, 0, hIcon)
                        End If
                    End Sub)
                Try
                    Using icon As System.Drawing.Icon = System.Drawing.Icon.FromHandle(hIcon)
                        Using bitmap = icon.ToBitmap()
                            Dim hBitmap As IntPtr
                            Try
                                hBitmap = bitmap.GetHbitmap()
                                Dim image As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                                image.Freeze()
                                SyncLock _overlayIconLock
                                    If Not _icons2.ContainsKey(String.Format("{0}_{1}", index, size)) Then
                                        _icons2.Add(String.Format("{0}_{1}", index, size), image)
                                    End If
                                End SyncLock
                            Finally
                                If Not IntPtr.Zero.Equals(hBitmap) Then
                                    Functions.DeleteObject(hBitmap)
                                    hBitmap = IntPtr.Zero
                                End If
                            End Try
                        End Using
                    End Using
                Finally
                    If Not IntPtr.Zero.Equals(hIcon) Then
                        Functions.DestroyIcon(hIcon)
                        hIcon = IntPtr.Zero
                    End If
                End Try
            End If

            Return _icons2(String.Format("{0}_{1}", index, size))
        End Function

        Public Shared Function ExtractIcon(ref As String, isSmall As Boolean) As BitmapSource
            If Not _icons.ContainsKey(ref.ToLower().Trim()) Then
                Dim s() As String = Split(ref, ","), icon As IntPtr, iconl As IntPtr
                Try
                    Functions.ExtractIconEx(s(0), Convert.ToInt32(s(1), CultureInfo.InvariantCulture), iconl, icon, 1)
                    If Not IntPtr.Zero.Equals(icon) Then
                        Dim img As BitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHIcon(If(isSmall, icon, iconl), Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
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
                        icon = IntPtr.Zero
                    End If
                    If Not IntPtr.Zero.Equals(iconl) Then
                        Functions.DestroyIcon(iconl)
                        iconl = IntPtr.Zero
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
End Namespace