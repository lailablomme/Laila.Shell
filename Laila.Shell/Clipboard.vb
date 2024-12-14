Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Public Class Clipboard
    Public Shared Function CanCopy(items As IEnumerable(Of Item)) As Boolean
        Return Not items Is Nothing AndAlso items.Count > 0 AndAlso items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANCOPY))
    End Function

    Public Shared Function CanCut(items As IEnumerable(Of Item)) As Boolean
        Return Not items Is Nothing AndAlso items.Count > 0 AndAlso items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANMOVE))
    End Function

    Public Shared Function CanPaste(folder As Folder) As Boolean
        ' check for paste by checking if it would accept a drop
        Dim dataObject As IDataObject
        Functions.OleGetClipboard(dataObject)

        Dim dropTarget As IDropTarget, dropTargetPtr As IntPtr, shellFolder As IShellFolder
        Try
            If Not folder.Parent Is Nothing Then
                shellFolder = folder.Parent.ShellFolder
                shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.RelativePIDL}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
            Else
                ' desktop
                shellFolder = Shell.Desktop.ShellFolder
                shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
            End If
            If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                dropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
            Else
                dropTarget = Nothing
            End If

            If Not dropTarget Is Nothing Then
                Dim effect As DROPEFFECT = Laila.Shell.DROPEFFECT.DROPEFFECT_COPY
                Dim hr As HRESULT = dropTarget.DragEnter(dataObject, 0, New WIN32POINT(), effect)
                dropTarget.DragLeave()

                Return hr = HRESULT.S_OK AndAlso effect <> DROPEFFECT.DROPEFFECT_NONE
            End If
        Finally
            If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                Marshal.Release(dropTargetPtr)
            End If
            If Not dropTarget Is Nothing Then
                Marshal.ReleaseComObject(dropTarget)
            End If
            If Not dataObject Is Nothing Then
                Marshal.ReleaseComObject(dataObject)
            End If
        End Try

        Return False
    End Function

    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject

        dataObject = Clipboard.GetDataObjectFor(items(0).Parent, items)

        Functions.OleSetClipboard(dataObject)

        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetClipboard(DROPEFFECT.DROPEFFECT_COPY)
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject

        dataObject = Clipboard.GetDataObjectFor(items(0).Parent, items)

        Functions.OleSetClipboard(dataObject)

        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetClipboard(DROPEFFECT.DROPEFFECT_MOVE)
    End Sub

    Public Shared Sub PasteFiles(folder As Folder)
        Dim dataObject As IDataObject
        Functions.OleGetClipboard(dataObject)

        Dim dropTarget As IDropTarget, dropTargetPtr As IntPtr, shellFolder As IShellFolder
        Try
            If Not folder.Parent Is Nothing Then
                shellFolder = folder.Parent.ShellFolder
                shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.RelativePIDL}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
            Else
                ' desktop
                shellFolder = Shell.Desktop.ShellFolder
                shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTargetPtr)
            End If
            If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                dropTarget = Marshal.GetTypedObjectForIUnknown(dropTargetPtr, GetType(IDropTarget))
            Else
                dropTarget = Nothing
            End If

            If Not dropTarget Is Nothing Then
                Dim effect As DROPEFFECT = ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.GetClipboard()
                Dim grfKeyState As MK = MK.MK_LBUTTON
                If effect = DROPEFFECT.DROPEFFECT_COPY Then grfKeyState = grfKeyState Or MK.MK_CONTROL
                Dim hr As HRESULT = dropTarget.DragEnter(dataObject, grfKeyState, New WIN32POINT(), effect)
                dropTarget.Drop(dataObject, grfKeyState, New WIN32POINT(), effect)
            End If
        Finally
            If Not IntPtr.Zero.Equals(dropTargetPtr) Then
                Marshal.Release(dropTargetPtr)
            End If
            If Not dropTarget Is Nothing Then
                Marshal.ReleaseComObject(dropTarget)
            End If
            If Not dataObject Is Nothing Then
                Marshal.ReleaseComObject(dataObject)
            End If
        End Try
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
        Try
            Functions.OleGetClipboard(dataObject)
            Return GetHasGlobalData(dataObject, clipboardFormat)
        Finally
            If Not dataObject Is Nothing Then
                Marshal.ReleaseComObject(dataObject)
            End If
        End Try
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

    Public Shared Function GetDataObjectFor(folder As Folder, items As IEnumerable(Of Item)) As IDataObject
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
        End Try

        Return result
    End Function
End Class
