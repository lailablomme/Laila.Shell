Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows.Forms

Public Class Clipboard
    Private Const CFSTR_SHELLIDLIST As String = "Shell IDList Array"

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

    Public Shared Function GetFileNameList(dataObj As ComTypes.IDataObject) As String()
        Dim files() As String
        files = readCFSTR_SHELLIDLIST(dataObj)?.Select(Function(i) i.FullPath).ToArray()
        If files Is Nothing OrElse files.Count = 0 Then
            files = readCF_HDROP(dataObj)
        End If
        Return files
    End Function

    Private Shared Function readCFSTR_SHELLIDLIST(dataObj As ComTypes.IDataObject) As List(Of Item)
        Try
            Dim format As New FORMATETC With {
                .cfFormat = Functions.RegisterClipboardFormat("Shell IDList Array"),
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            If dataObj.QueryGetData(format) = 0 Then
                Dim medium As STGMEDIUM
                dataObj.GetData(format, medium)

                Return Pidl.GetItemsFromShellIDListArray(medium.unionmember)
            Else
                Return Nothing
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Private Shared Function readCF_HDROP(dataObj As ComTypes.IDataObject) As String()
        Try
            Dim format As New FORMATETC With {
                .cfFormat = ClipboardFormat.CF_HDROP,
                .ptd = IntPtr.Zero,
                .dwAspect = DVASPECT.DVASPECT_CONTENT,
                .lindex = -1,
                .tymed = TYMED.TYMED_HGLOBAL
            }
            If dataObj.QueryGetData(format) = 0 Then
                Dim medium As STGMEDIUM
                dataObj.GetData(format, medium)

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

End Class
