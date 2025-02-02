Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Threading

Public Class Clipboard
    Public Shared Function CanCopy(items As IEnumerable(Of Item)) As Boolean
        Return Not items Is Nothing AndAlso items.Count > 0 AndAlso items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANCOPY))
    End Function

    Public Shared Function CanCut(items As IEnumerable(Of Item)) As Boolean
        Return Not items Is Nothing AndAlso items.Count > 0 AndAlso items.All(Function(i) i.Attributes.HasFlag(SFGAO.CANMOVE))
    End Function

    Public Shared Function CanPaste(folder As Folder) As Boolean
        Return Shell.RunOnSTAThread(
            Function() As Boolean
                ' check for paste by checking if it would accept a drop
                Dim dataObject As ComTypes.IDataObject
                Functions.OleGetClipboard(dataObject)

                Dim dropTarget As IDropTarget
                Try
                    If Not folder.Parent Is Nothing Then
                        folder.Parent.ShellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.RelativePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                    Else
                        ' desktop
                        Shell.Desktop.ShellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                    End If

                    If Not dropTarget Is Nothing Then
                        Dim effect As DROPEFFECT = Laila.Shell.DROPEFFECT.DROPEFFECT_COPY
                        Dim hr As HRESULT = dropTarget.DragEnter(dataObject, 0, New WIN32POINT(), effect)
                        dropTarget.DragLeave()

                        Return hr = HRESULT.S_OK AndAlso effect <> DROPEFFECT.DROPEFFECT_NONE
                    End If
                Finally
                    If Not dropTarget Is Nothing Then
                        Marshal.ReleaseComObject(dropTarget)
                        dropTarget = Nothing
                    End If
                    If Not dataObject Is Nothing Then
                        Marshal.ReleaseComObject(dataObject)
                        dataObject = Nothing
                    End If
                End Try

                Return False
            End Function)
    End Function

    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        Dim dataObject As ComTypes.IDataObject

        dataObject = Clipboard.GetDataObjectFor(items(0).Parent, items)
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_COPY)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        Dim dataObject As ComTypes.IDataObject

        dataObject = Clipboard.GetDataObjectFor(items(0).Parent, items)
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_MOVE)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Sub PasteFiles(folder As Folder)
        Dim thread As Thread = New Thread(New ThreadStart(
            Sub()
                Dim dataObject As ComTypes.IDataObject
                Functions.OleGetClipboard(dataObject)

                Dim dropTarget As IDropTarget, shellFolder As IShellFolder
                Try
                    If Not folder.Parent Is Nothing Then
                        shellFolder = folder.Parent.ShellFolder
                        shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.RelativePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                    Else
                        ' desktop
                        shellFolder = Shell.Desktop.ShellFolder
                        shellFolder.GetUIObjectOf(IntPtr.Zero, 1, {folder.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, dropTarget)
                    End If

                    If Not dropTarget Is Nothing Then
                        Dim effect As DROPEFFECT = ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.GetClipboard()
                        Dim grfKeyState As MK = MK.MK_LBUTTON
                        If effect = DROPEFFECT.DROPEFFECT_COPY Then grfKeyState = grfKeyState Or MK.MK_CONTROL
                        Dim hr As HRESULT = dropTarget.DragEnter(dataObject, grfKeyState, New WIN32POINT(), effect)
                        dropTarget.Drop(dataObject, grfKeyState, New WIN32POINT(), effect)
                    End If
                Finally
                    If Not dropTarget Is Nothing Then
                        Marshal.ReleaseComObject(dropTarget)
                        dropTarget = Nothing
                    End If
                    If Not dataObject Is Nothing Then
                        Marshal.ReleaseComObject(dataObject)
                        dataObject = Nothing
                    End If
                End Try
            End Sub))

        thread.SetApartmentState(ApartmentState.STA)
        thread.Start()
    End Sub

    Public Shared Function GetFileNameList(dataObj As ComTypes.IDataObject) As String()
        Dim files() As String
        files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(dataObj)?.Select(Function(i) i.FullPath).ToArray()
        If files Is Nothing OrElse files.Count = 0 Then
            files = ClipboardFormats.CF_HDROP.GetData(dataObj)
        End If
        Return files
    End Function

    Public Shared Function GetHasGlobalData(clipboardFormat As String)
        Dim dataObject As ComTypes.IDataObject
        Try
            Functions.OleGetClipboard(dataObject)
            Return GetHasGlobalData(dataObject, clipboardFormat)
        Finally
            If Not dataObject Is Nothing Then
                Marshal.ReleaseComObject(dataObject)
                dataObject = Nothing
            End If
        End Try
    End Function

    Public Shared Function GetHasGlobalData(dataObject As ComTypes.IDataObject, clipboardFormat As String)
        Return GetHasGlobalData(dataObject, Functions.RegisterClipboardFormat(clipboardFormat))
    End Function

    Public Shared Function GetHasGlobalData(dataObject As ComTypes.IDataObject, clipboardFormat As Short)
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = clipboardFormat,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Return dataObject.QueryGetData(format) = 0
    End Function

    Public Shared Function GetDataObjectFor(folder As Folder, items As IEnumerable(Of Item)) As ComTypes.IDataObject
        Dim result As ComTypes.IDataObject

        ' make a DataObject for our list of items
        Dim inner As DataObject = New DataObject()
        Functions.SHCreateDataObject(folder.Pidl.Clone().AbsolutePIDL, items.Count,
            items.Select(Function(i) i.Pidl.Clone().RelativePIDL).ToArray(),
            Nothing, GetType(IDataObject).GUID, result)
        'result = inner
        'ClipboardFormats.CF_HDROP.SetData(result, items)
        'ClipboardFormats.CFSTR_FILEDESCRIPTOR.SetData(result, items)
        'ClipboardFormats.CFSTR_SHELLIDLIST.SetData(result, items)
        'ClipboardFormats.CFSTR_FILECONTENTS.SetData(result, items)

        ' for some reason we can't properly write to our DataObject before a DropTarget initializes it,
        ' and I don't know what it's doing 
        Dim initDropTarget As IDropTarget
        Try
            Shell.Desktop.ShellFolder.GetUIObjectOf(IntPtr.Zero, 1, {Shell.Desktop.Pidl.AbsolutePIDL}, GetType(IDropTarget).GUID, 0, initDropTarget)
            initDropTarget.DragEnter(result, 0, New WIN32POINT() With {.x = 0, .y = 0}, 0)
            initDropTarget.DragLeave()
        Finally
            If Not initDropTarget Is Nothing Then
                Marshal.ReleaseComObject(initDropTarget)
                initDropTarget = Nothing
            End If
        End Try

        Return result
    End Function

    Public Shared Function GetGlobalDataDWord(dataObject As ComTypes.IDataObject, clipboardFormat As String) As Integer?
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = Functions.RegisterClipboardFormat(clipboardFormat),
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Dim medium As STGMEDIUM
        Dim result As Integer = dataObject.QueryGetData(format)
        If result = 0 Then
            dataObject.GetData(format, medium)
            Return Marshal.ReadInt32(medium.unionmember)
        Else
            Return Nothing
        End If
    End Function

    Public Shared Sub SetGlobalDataDWord(dataObject As ComTypes.IDataObject, clipboardFormat As String, val As Integer)
        Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(Of Integer))
        Marshal.WriteInt32(ptr, val)
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = Functions.RegisterClipboardFormat(clipboardFormat),
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Dim medium As STGMEDIUM = New STGMEDIUM() With {
            .pUnkForRelease = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL,
            .unionmember = ptr
        }
        dataObject.SetData(format, medium, True)
    End Sub
End Class
