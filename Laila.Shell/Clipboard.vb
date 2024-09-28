Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows.Forms

Public Class Clipboard
    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        setData(items.Where(Function(i) i.Attributes.HasFlag(SFGAO.CANCOPY)))
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        setData(items.Where(Function(i) i.Attributes.HasFlag(SFGAO.CANMOVE)))

        For Each item In items
            item.IsCut = True
        Next

        Dim effect As Integer = DROPEFFECT.DROPEFFECT_MOVE
        Dim globalPtr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(effect))
        Marshal.StructureToPtr(effect, globalPtr, False)

        Functions.OpenClipboard(IntPtr.Zero)
        Functions.SetClipboardData(Functions.RegisterClipboardFormat(Functions.CFSTR_PREFERREDDROPEFFECT), globalPtr)
        Functions.CloseClipboard()
    End Sub

    Private Shared Sub setData(items As IEnumerable(Of Item))
        Dim dropFiles As New DROPFILES()
        dropFiles.pFiles = Marshal.SizeOf(dropFiles)
        dropFiles.fWide = False

        Dim fileNames As New StringBuilder()
        For Each item In items
            fileNames.Append(item.FullPath)
            fileNames.Append(Chr(0)) ' Null character
        Next
        fileNames.Append(Chr(0)) ' Double null character to end the list

        Dim fileNamesBytes As Byte() = Encoding.ASCII.GetBytes(fileNames.ToString())
        Dim globalPtr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(dropFiles) + fileNamesBytes.Length)
        Marshal.StructureToPtr(dropFiles, globalPtr, False)
        Marshal.Copy(fileNamesBytes, 0, IntPtr.Add(globalPtr, Marshal.SizeOf(dropFiles)), fileNamesBytes.Length)

        Functions.OpenClipboard(IntPtr.Zero)
        Try
            Functions.EmptyClipboard()
            Functions.SetClipboardData(ClipboardFormat.CF_HDROP, globalPtr).ToString()
        Finally
            Functions.CloseClipboard()
        End Try
    End Sub
End Class
