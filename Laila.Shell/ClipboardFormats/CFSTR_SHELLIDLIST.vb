Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization

Namespace ClipboardFormats
    Public Class CFSTR_SHELLIDLIST
        Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"

        Public Shared Function GetData(dataObject As IDataObject) As List(Of Item)
            Try
                Dim format As New FORMATETC With {
                    .cfFormat = Functions.RegisterClipboardFormat(CFSTR_SHELLIDLIST),
                    .ptd = IntPtr.Zero,
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .lindex = -1,
                    .tymed = TYMED.TYMED_HGLOBAL
                }
                If dataObject.QueryGetData(format) = 0 Then
                    Dim medium As STGMEDIUM
                    dataObject.GetData(format, medium)
                    Dim ptr As IntPtr = medium.unionmember
                    Dim test As Integer = Marshal.ReadInt32(medium.unionmember)
                    If test >= &H10000 OrElse test <= -1 Then
                        ' windows explorer gives a pointer to a pointer to a shellidlist,
                        ' however SHCreateDataObject does not
                        ptr = Marshal.ReadIntPtr(medium.unionmember)
                    End If
                    Return Pidl.GetItemsFromShellIDListArray(ptr)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Sub SetData(dataObject As IDataObject, items As IEnumerable(Of Item))
            Dim ptr As IntPtr = Pidl.CreateShellIDListArray(items)
            Dim format As FORMATETC = New FORMATETC With {
                .cfFormat = Functions.RegisterClipboardFormat(CFSTR_SHELLIDLIST),
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            Dim medium As STGMEDIUM = New STGMEDIUM With {
                .tymed = TYMED.TYMED_HGLOBAL,
                .unionmember = ptr,
                .pUnkForRelease = IntPtr.Zero
            }
            dataObject.SetData(format, medium, True)
        End Sub
    End Class
End Namespace