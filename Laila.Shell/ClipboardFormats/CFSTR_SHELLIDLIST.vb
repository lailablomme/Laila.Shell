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

                    Return Pidl.GetItemsFromShellIDListArray(medium.unionmember)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Sub SetData(dataObject As IDataObject, items As IEnumerable(Of Item))
            Dim format As FORMATETC = New FORMATETC With {
                .cfFormat = Functions.RegisterClipboardFormat(CFSTR_SHELLIDLIST),
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            Dim medium As STGMEDIUM = New STGMEDIUM With {
                .tymed = TYMED.TYMED_HGLOBAL,
                .unionmember = Pidl.CreateShellIDListArray(items),
                .pUnkForRelease = IntPtr.Zero
            }
            dataObject.SetData(format, medium, True)
        End Sub
    End Class
End Namespace