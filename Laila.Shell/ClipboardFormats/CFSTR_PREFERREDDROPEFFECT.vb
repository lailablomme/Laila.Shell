Imports System.CodeDom
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Windows.Media.Animation

Namespace ClipboardFormats
    Public Class CFSTR_PREFERREDDROPEFFECT
        Private Const CFSTR_PREFERREDDROPEFFECT As String = "Preferred DropEffect"

        Public Shared Sub SetData(dataObject As IDataObject, preferredDropEffect As DROPEFFECT)
            Dim effect As Byte = If(preferredDropEffect = DROPEFFECT.DROPEFFECT_MOVE, 2, 5)
            Dim hGlobal As IntPtr = Marshal.AllocHGlobal(4)
            Marshal.WriteByte(hGlobal, 0, effect)

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

        Public Shared Sub SetClipboard(preferredDropEffect As DROPEFFECT)
            Dim effect As Integer = If(preferredDropEffect = DROPEFFECT.DROPEFFECT_COPY, 5, 2)
            Dim hGlobal As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of Integer))
            Marshal.WriteInt32(hGlobal, effect)

            Dim format As UInteger = Functions.RegisterClipboardFormatWIN32(CFSTR_PREFERREDDROPEFFECT)
            Functions.OpenClipboard(IntPtr.Zero)
            Functions.SetClipboardData(format, hGlobal)
            Functions.CloseClipboard()
        End Sub

        Public Shared Function GetClipboard() As DROPEFFECT
            Dim format As UInteger = Functions.RegisterClipboardFormatWIN32(CFSTR_PREFERREDDROPEFFECT)
            Functions.OpenClipboard(IntPtr.Zero)
            Dim hGlobal As IntPtr = Functions.GetClipboardData(format)
            Try
                Return If(IntPtr.Zero.Equals(hGlobal), DROPEFFECT.DROPEFFECT_COPY,
                    If(Marshal.ReadInt32(hGlobal) = 5, DROPEFFECT.DROPEFFECT_COPY, DROPEFFECT.DROPEFFECT_MOVE))
            Finally
                Functions.CloseClipboard()
            End Try
        End Function
    End Class
End Namespace