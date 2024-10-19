Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows.Forms

Public Class Clipboard
    Private Shared ReadOnly GMEM_MOVEABLE As Integer = &H2

    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject = New DragDataObject()

        ClipboardFormats.CFSTR_SHELLIDLIST.SetData(dataObject, items)
        ClipboardFormats.CFSTR_FILEDESCRIPTOR.SetData(dataObject, items)
        ClipboardFormats.CFSTR_FILECONTENTS.SetData(dataObject, items)
        ClipboardFormats.CF_HDROP.SetData(dataObject, items)
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_COPY)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject = New DragDataObject()

        ClipboardFormats.CFSTR_SHELLIDLIST.SetData(dataObject, items)
        ClipboardFormats.CFSTR_FILEDESCRIPTOR.SetData(dataObject, items)
        ClipboardFormats.CFSTR_FILECONTENTS.SetData(dataObject, items)
        ClipboardFormats.CF_HDROP.SetData(dataObject, items)
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_MOVE)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Function GetFileNameList(dataObj As IDataObject) As String()
        Dim files() As String
        files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(dataObj)?.Select(Function(i) i.FullPath).ToArray()
        If files Is Nothing OrElse files.Count = 0 Then
            files = ClipboardFormats.CF_HDROP.GetData(dataObj)
        End If
        Return files
    End Function
End Class
