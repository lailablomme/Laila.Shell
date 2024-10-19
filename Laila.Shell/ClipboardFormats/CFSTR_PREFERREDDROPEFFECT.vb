Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization

Namespace ClipboardFormats
    Public Class CFSTR_PREFERREDDROPEFFECT
        Private Const CFSTR_PREFERREDDROPEFFECT As String = "Preferred DropEffect"

        Public Shared Sub SetData(dataObject As IDataObject, preferredDropEffect As DROPEFFECT)
            Dim effect As Integer = DROPEFFECT.DROPEFFECT_MOVE
            Dim hGlobal As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(effect))
            Marshal.StructureToPtr(effect, hGlobal, False)

            Dim format As FORMATETC = New FORMATETC With {
                .cfFormat = Functions.RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT),
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            Dim medium As STGMEDIUM = New STGMEDIUM With {
                .tymed = TYMED.TYMED_HGLOBAL,
                .unionmember = hGlobal,
                .pUnkForRelease = IntPtr.Zero
            }
            dataObject.SetData(format, medium, True)
        End Sub
    End Class
End Namespace