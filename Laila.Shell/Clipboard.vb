Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows.Forms

Public Class Clipboard
    Private Shared ReadOnly GMEM_MOVEABLE As Integer = &H2

    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject

        Using parent = items(0).GetParent()
            dataObject = Clipboard.GetDataObjectFor(parent, items)
        End Using

        Functions.OleSetClipboard(dataObject)

        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetClipboard(DROPEFFECT.DROPEFFECT_COPY)
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject

        Using parent = items(0).GetParent()
            dataObject = Clipboard.GetDataObjectFor(parent, items)
        End Using

        Functions.OleSetClipboard(dataObject)

        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetClipboard(DROPEFFECT.DROPEFFECT_MOVE)
    End Sub

    Public Shared Function GetFileNameList(dataObj As IDataObject) As String()
        Dim files() As String
        files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(dataObj)?.Select(Function(i) i.FullPath).ToArray()
        If files Is Nothing OrElse files.Count = 0 Then
            files = ClipboardFormats.CF_HDROP.GetData(dataObj)
        End If
        Return files
    End Function

    Public Shared Function GetHasGlobalData(clipboardFormat As String)
        Dim dataObject As IDataObject
        Functions.OleGetClipboard(dataObject)
        Return GetHasGlobalData(dataObject, clipboardFormat)
    End Function

    Public Shared Function GetHasGlobalData(dataObject As IDataObject, clipboardFormat As String)
        Return GetHasGlobalData(dataObject, Functions.RegisterClipboardFormat(clipboardFormat))
    End Function

    Public Shared Function GetHasGlobalData(dataObject As IDataObject, clipboardFormat As Short)
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = clipboardFormat,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Return dataObject.QueryGetData(format) = 0
    End Function

    Public Shared Function GetDataObjectFor(folder As Folder, items As List(Of Item)) As IDataObject
        Dim result As IDataObject

        ' make a DataObject for our list of items
        Functions.SHCreateDataObject(folder.Pidl.Clone().AbsolutePIDL, items.Count,
            items.Select(Function(i) i.Pidl.Clone().RelativePIDL).ToArray(), IntPtr.Zero, GetType(IDataObject).GUID, result)

        ' for some reason we can't properly write to our DataObject before a DropTarget initializes it,
        ' and I don't know what it's doing 
        Dim initDropTarget As IDropTarget, pidl As IntPtr, dropTargetPtr As IntPtr
        Dim shellFolder As IShellFolder
        Try
            Functions.SHGetDesktopFolder(shellFolder)
            shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {Shell.Desktop.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
            initDropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
            initDropTarget.DragEnter(result, 0, New WIN32POINT() With {.x = 0, .y = 0}, 0)
            initDropTarget.DragLeave()
        Finally
            If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                Marshal.Release(dropTargetPtr)
            End If
            If Not IntPtr.Zero.Equals(pidl) Then
                Marshal.FreeCoTaskMem(pidl)
            End If
            If Not initDropTarget Is Nothing Then
                Marshal.ReleaseComObject(initDropTarget)
            End If
            If Not shellFolder Is Nothing Then
                Marshal.ReleaseComObject(shellFolder)
            End If
        End Try

        Return result
    End Function
End Class
