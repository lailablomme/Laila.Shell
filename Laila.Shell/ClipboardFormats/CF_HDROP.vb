Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text

Namespace ClipboardFormats
    Public Class CF_HDROP
        Private Shared ReadOnly GMEM_MOVEABLE As Integer = &H2

        Public Shared Function GetData(dataObject As IDataObject) As String()
            Try
                Dim format As New FORMATETC With {
                    .cfFormat = ClipboardFormat.CF_HDROP,
                    .ptd = IntPtr.Zero,
                    .dwAspect = DVASPECT.DVASPECT_CONTENT,
                    .lindex = -1,
                    .tymed = TYMED.TYMED_HGLOBAL
                }
                If dataObject.QueryGetData(format) = 0 Then
                    Dim medium As STGMEDIUM
                    dataObject.GetData(format, medium)

                    Dim fileCount As UInteger = Functions.DragQueryFile(medium.unionmember, UInt32.MaxValue, Nothing, 0)
                    Dim fileList As New List(Of String)()
                    If fileCount > 0 Then
                        For i As UInteger = 0 To fileCount - 1
                            Dim filePathBuilder As New StringBuilder(260) ' MAX_PATH
                            Functions.DragQueryFile(medium.unionmember, i, filePathBuilder, CType(filePathBuilder.Capacity, UInteger))
                            fileList.Add(filePathBuilder.ToString())
                        Next
                    End If

                    Return If(fileList.Count > 0, fileList.ToArray(), Nothing)
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Return Nothing
            End Try
        End Function

        Public Shared Sub SetData(dataObject As IDataObject, items As IEnumerable(Of Item))
            Dim sb As StringBuilder = New StringBuilder()
            For Each filePath As String In items.Select(Function(i) i.FullPath).ToArray()
                sb.Append(filePath & Chr(0))
            Next
            sb.Append(Chr(0))

            Dim strPtr As IntPtr = Marshal.StringToHGlobalUni(sb.ToString())
            Dim strSize As Integer = Functions.GlobalSize(strPtr)

            Dim hGlobal As IntPtr = Functions.GlobalAlloc(GMEM_MOVEABLE, Marshal.SizeOf(Of DROPFILES) + strSize)
            Dim pGlobal As IntPtr = Functions.GlobalLock(hGlobal)

            ' Write the DROPFILES structure
            Dim dropfiles As New DROPFILES()
            dropfiles.pFiles = CType(Marshal.SizeOf(GetType(DROPFILES)), UInteger)
            dropfiles.fWide = True ' Using Unicode (WCHAR)

            Marshal.StructureToPtr(dropfiles, pGlobal, False)
            Functions.RtlMoveMemory(IntPtr.Add(pGlobal, dropfiles.pFiles), strPtr, strSize)

            If Not IntPtr.Zero.Equals(strPtr) Then
                Marshal.FreeHGlobal(strPtr)
                strPtr = IntPtr.Zero
            End If
            Functions.GlobalUnlock(pGlobal)

            Dim format As FORMATETC = New FORMATETC With {
                .cfFormat = ClipboardFormat.CF_HDROP,
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