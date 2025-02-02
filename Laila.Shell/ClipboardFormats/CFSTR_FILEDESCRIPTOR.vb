Imports System.IO
Imports System.Runtime
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace ClipboardFormats
    Public Class CFSTR_FILEDESCRIPTOR
        Public Shared Sub SetData(dataObject As ComTypes.IDataObject, items As IEnumerable(Of Item))
            Dim filePaths() As String = items.Select(Function(i) i.FullPath).ToArray()

            ' Create the file group descriptor for the number of files being dragged
            Dim fileDescriptors As List(Of FILEDESCRIPTORW) = New List(Of FILEDESCRIPTORW)()
            For Each item In items
                addFileDescriptors(item, fileDescriptors)
            Next
            Dim cItems As UInt32 = CUInt(fileDescriptors.Count)

            Dim ptr As IntPtr = Marshal.AllocHGlobal(
                CInt(Marshal.SizeOf(cItems) + Marshal.SizeOf(Of FILEDESCRIPTORW) * cItems))
            Marshal.WriteInt32(ptr, cItems)
            For i = 0 To cItems - 1
                Marshal.StructureToPtr(
                    fileDescriptors(i),
                    IntPtr.Add(ptr, Marshal.SizeOf(cItems) + i * Marshal.SizeOf(Of FILEDESCRIPTORW)),
                    False)
            Next

            ' Set the "FileGroupDescriptorW" data using COM IDataObject
            Dim format As New FORMATETC With {
                .cfFormat = Functions.RegisterClipboardFormat("FileGroupDescriptorW"),
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .tymed = TYMED.TYMED_HGLOBAL,
                .lindex = -1,
                .ptd = IntPtr.Zero
            }

            Dim medium As New STGMEDIUM With {
                .tymed = TYMED.TYMED_HGLOBAL,
                .unionmember = ptr,
                .pUnkForRelease = IntPtr.Zero
            }

            dataObject.SetData(format, medium, True)
            'Functions.OpenClipboard(IntPtr.Zero)
            '  Functions.EmptyClipboard()
            '  Functions.SetClipboardData(Functions.RegisterClipboardFormatWIN32("FileGroupDescriptorW"), ptr).ToString()
            '  Functions.CloseClipboard()

        End Sub

        Private Shared Sub addFileDescriptors(item As Item, fileDescriptors As List(Of FILEDESCRIPTORW), Optional root As String = "")
            Dim data As WIN32_FILE_ATTRIBUTE_DATA
            Functions.GetFileAttributesEx(item.FullPath, 0, data)

            Dim name As String = IO.Path.Combine(root, item.FullPath.Substring(item.FullPath.LastIndexOf(IO.Path.DirectorySeparatorChar) + 1))

            ' Populate FILEDESCRIPTORW for each file
            Dim fdesc As FILEDESCRIPTORW = New FILEDESCRIPTORW With {
                .dwFlags = FD.FD_FILESIZE Or FD.FD_UNICODE Or FD.FD_ATTRIBUTES Or FD.FD_ACCESSTIME Or FD.FD_CREATETIME Or FD.FD_WRITESTIME,
                .cFileName = name,
                .dwFileAttributes = data.dwFileAttributes,
                .nFileSizeLow = data.nFileSizeLow,
                .nFileSizeHigh = data.nFileSizeHigh,
                .ftCreationTime = data.ftCreationTime,
                .ftLastAccessTime = data.ftLastAccessTime,
                .ftLastWriteTime = data.ftLastWriteTime
            }

            fileDescriptors.Add(fdesc)

            If data.dwFileAttributes.HasFlag(FILE_ATTRIBUTE.FILE_ATTRIBUTE_DIRECTORY) Then
                For Each subItem In CType(item, Folder).GetItems()
                    addFileDescriptors(subItem, fileDescriptors, name)
                Next
            End If
        End Sub
    End Class
End Namespace