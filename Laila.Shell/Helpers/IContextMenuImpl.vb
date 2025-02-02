Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Text
Imports System.Windows.Media.Animation
Imports Microsoft.Win32

Namespace Helpers
    Public Class IContextMenuImpl
        Implements IContextMenu, IShellExtInit, IDisposable

        Public Property SelectedItems As IEnumerable(Of Item)

        Private _hmenu As IntPtr, _iMenu As Integer, _idCmdFirst As Integer, _idCmdLast As Integer, _uFlags As Integer
        Private disposedValue As Boolean
        Private _handlers As Dictionary(Of Guid, Tuple(Of IContextMenu, IDataObject)) = New Dictionary(Of Guid, Tuple(Of IContextMenu, IDataObject))()

        Public Function QueryContextMenu(hmenu As IntPtr, iMenu As Integer, idCmdFirst As Integer, idCmdLast As Integer, uFlags As Integer) As Integer Implements IContextMenu.QueryContextMenu
            _hmenu = hmenu
            _iMenu = iMenu
            _idCmdFirst = idCmdFirst
            _idCmdLast = idCmdLast
            _uFlags = uFlags
            'For Each item In items
            '    If item.FullPath.StartsWith("::{") Then
            '        Dim clsid As Guid = New Guid(item.FullPath.Substring(3, item.FullPath.Length - 4))

            '    Else

            '    End If
            'Next
        End Function

        Public Function InvokeCommand(ByRef pici As CMInvokeCommandInfoEx) As Integer Implements IContextMenu.InvokeCommand
            Return HRESULT.S_OK
        End Function

        Public Function GetCommandString(idcmd As Integer, uflags As Integer, reserved As Integer, commandstring As StringBuilder, cch As Integer) As Integer Implements IContextMenu.GetCommandString
            Return HRESULT.S_OK
        End Function

        Public Function Initialize(pidlFolder As IntPtr, lpdobj As IDataObject, hKeyProgID As IntPtr) As Integer Implements IShellExtInit.Initialize
            Dim items As List(Of Item) = ClipboardFormats.CF_HDROP.GetData(lpdobj)?.ToList().Select(Function(s) Item.FromParsingName(s, Nothing)).ToList()

            ' enumerate all context menu handlers
            Dim cmhandlers As Dictionary(Of String, Guid) = New Dictionary(Of String, Guid)()
            addHandlersFromKey("*", cmhandlers)
            If Not items Is Nothing Then
                For Each item In items
                    addHandlersForItem(item, cmhandlers)
                Next
            End If

            ' delegate querying context menu items to each handler
            For Each cmhandler In cmhandlers
                Try
                    Dim handler As IContextMenu = Activator.CreateInstance(Type.GetTypeFromCLSID(cmhandler.Value))

                    Dim dobj As IDataObject = New DataObject()
                    ClipboardFormats.CF_HDROP.SetData(dobj, Me.SelectedItems)
                    Dim shellExtInit As IShellExtInit = TryCast(handler, IShellExtInit)
                    If Not shellExtInit Is Nothing Then
                        shellExtInit.Initialize(Shell.Desktop.Pidl.AbsolutePIDL, dobj, hKeyProgID)
                    End If

                    handler.QueryContextMenu(_hmenu, _iMenu, _idCmdFirst, _idCmdLast, _uFlags)

                    _handlers.Add(cmhandler.Value, New Tuple(Of IContextMenu, IDataObject)(handler, dobj))

                    For i = 0 To Functions.GetMenuItemCount(_hmenu) - 1
                        Dim mii As MENUITEMINFO = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_ID
                        Functions.GetMenuItemInfo(_hmenu, i, True, mii)
                        If mii.wID > _idCmdFirst Then
                            _idCmdFirst = mii.wID
                        End If
                    Next
                    _idCmdFirst += 1
                Catch ex As Exception
                End Try
            Next

            Return HRESULT.S_OK
        End Function

        Private Sub addHandlersForItem(item As Item, handlers As Dictionary(Of String, Guid))
            If TypeOf item Is Folder Then
                addHandlersFromKey("Folder", handlers)
            ElseIf item.FullPath?.StartsWith("::{") Then
                Dim clsid As Guid = New Guid(item.FullPath.Substring(3, item.FullPath.Length - 4))
                addHandlersFromKey("CLSID\" & clsid.ToString(), handlers)
            Else
                Dim ext As String = IO.Path.GetExtension(item.FullPath)
                If ext?.Length > 1 Then
                    addHandlersFromKey(ext, handlers)
                    Using key = Registry.ClassesRoot.OpenSubKey(ext)
                        Dim className As String = key.GetValue("").ToString()
                        If Not String.IsNullOrWhiteSpace(className) Then
                            addHandlersFromKey(className, handlers)
                        End If
                    End Using
                End If
            End If
        End Sub

        Private Sub addHandlersFromKey(keyName As String, handlers As Dictionary(Of String, Guid))
            Try
                Using key = Registry.ClassesRoot.OpenSubKey(keyName & "\shellex\ContextMenuHandlers")
                    For Each subKeyName In key.GetSubKeyNames()
                        Using subKey = Registry.ClassesRoot.OpenSubKey(keyName & "\shellex\ContextMenuHandlers\" & subKeyName)
                            Dim handlerName As String, handlerGuid As Guid
                            If subKeyName.StartsWith("{") AndAlso subKeyName.EndsWith("}") Then
                                handlerName = keyName & "\" & subKey.GetValue("").ToString()
                                handlerGuid = New Guid(subKeyName)
                            Else
                                handlerName = keyName & "\" & subKeyName
                                handlerGuid = New Guid(subKey.GetValue("").ToString())
                            End If
                            If Not handlers.ContainsKey(handlerName) Then
                                handlers.Add(handlerName, handlerGuid)
                            End If
                        End Using
                    Next
                End Using
            Catch ex As Exception
            End Try
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True

                For Each handler In _handlers
                    Marshal.ReleaseComObject(handler.Value.Item1)
                    CType(handler.Value.Item2, IDisposable).Dispose()
                Next
                _handlers.Clear()
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace