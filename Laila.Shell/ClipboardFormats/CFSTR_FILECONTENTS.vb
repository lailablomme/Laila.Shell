Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports Laila.Shell.Helpers

Namespace ClipboardFormats
    Public Class CFSTR_FILECONTENTS
        Public Shared Sub SetData(dataObject As ComTypes.IDataObject, items As IEnumerable(Of Item))
            Dim index As Integer = -1
            For Each item In items
                setData(dataObject, item, index)
            Next
        End Sub

        Private Shared Sub setData(dataObject As ComTypes.IDataObject, item As Item, ByRef index As Integer)
            index += 1

            Dim attr As FILE_ATTRIBUTE = Functions.GetFileAttributesW(item.FullPath)

            If Not attr.HasFlag(FILE_ATTRIBUTE.FILE_ATTRIBUTE_DIRECTORY) Then
                ' Create an IStream for the data
                Dim stream As IStream = New StreamAdapter(item.FullPath)
                Dim ptr As IntPtr = Marshal.GetComInterfaceForObject(stream, GetType(IStream))

                ' Set the "FileContents" format using COM IDataObject
                Dim format As New FORMATETC With {
                    .cfFormat = Functions.RegisterClipboardFormat("FileContents"),
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .tymed = TYMED.TYMED_ISTREAM,
                    .lindex = index,
                    .ptd = IntPtr.Zero
                }

                ' Create the STGMEDIUM structure to pass the IStream
                Dim medium As New STGMEDIUM With {
                    .tymed = TYMED.TYMED_ISTREAM,
                    .unionmember = ptr,
                    .pUnkForRelease = IntPtr.Zero
                }

                ' Set the data on the IDataObject
                dataObject.SetData(format, medium, True)
                'Functions.OpenClipboard(IntPtr.Zero)
                'Functions.SetClipboardData(Functions.RegisterClipboardFormatWIN32("FileContents"), ptr).ToString()
                'Functions.CloseClipboard()
            Else
                For Each subItem In CType(item, Folder).GetItems()
                    setData(dataObject, subItem, index)
                Next
            End If
        End Sub
    End Class
End Namespace