Imports System.Collections.Concurrent
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Forms.PropertyGridInternal
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Application
Imports Laila.Shell.Interop.ContextMenu
Imports Laila.Shell.Interop.DragDrop
Imports Laila.Shell.Interop.Folders
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Windows
Imports Microsoft.Win32

Namespace Controls
    Public MustInherit Class BaseContextMenu
        Inherits BaseMenu

        Private _contextMenu As IContextMenu
        Private _contextMenu2 As IContextMenu2
        Private _contextMenu3 As IContextMenu3
        Protected _hMenu As IntPtr

        Protected Overrides Function DoRenameAfter(Tag As Tuple(Of Integer, String, Object)) As Boolean
            Dim newMenuItem As MenuItem = Me.Items.Cast(Of Control) _
                .FirstOrDefault(Function(c) TypeOf c Is MenuItem _
                    AndAlso Not c.Tag Is Nothing _
                    AndAlso CType(c.Tag, Tuple(Of Integer, String, Object)).Item2 = "New")

            Return Not newMenuItem Is Nothing AndAlso
                Not newMenuItem.Items.Cast(Of Control).FirstOrDefault(
                    Function(c) _invokedId.Equals(c.Tag)) Is Nothing
        End Function

        Public Function MakeButton(tag As Tuple(Of Integer, String, Object), toolTip As String) As Button
            Dim button As Button = New Button()
            Dim image As Image = New Image()
            image.Width = 16
            image.Height = 16
            image.Margin = New Thickness(2)
            Select Case tag.Item2
                Case "copy" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/copy16.png")
                Case "cut" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/cut16.png")
                Case "paste" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/paste16.png")
                Case "rename" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/rename16.png")
                Case "delete" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/delete16.png")
            End Select
            button.Content = image
            button.ToolTip = toolTip
            button.Tag = tag
            Return button
        End Function

        Public Function MakeToggleButton(tag As Tuple(Of Integer, String, Object), toolTip As String, isChecked As Boolean) As ToggleButton
            Dim button As ToggleButton = New ToggleButton()
            Dim image As Image = New Image()
            image.Width = 16
            image.Height = 16
            image.Margin = New Thickness(2)
            Select Case tag.Item2
                Case "laila.shell.(un)pin" : image.Source = New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/pin16.png")
            End Select
            button.Content = image
            button.ToolTip = toolTip
            button.Tag = tag
            button.IsChecked = isChecked
            Return button
        End Function

        Public Function GetMenuItems() As List(Of Control)
            Dim tcs4 As New TaskCompletionSource(Of List(Of MenuItemData))
            _thread.Run(
                Sub()
                    tcs4.SetResult(getMenuItemData(_hMenu, -1))
                End Sub)
            Return getMenuItems(tcs4.Task.Result)
        End Function

        Private Function getMenuItems(data As List(Of MenuItemData)) As List(Of Control)
            Dim result As List(Of Control) = New List(Of Control)()

            For Each item In data
                If item.Header = "-----" Then
                    result.Add(New Separator())
                Else
                    Dim menuItem As MenuItem = New MenuItem() With {
                        .Header = item.Header,
                        .Icon = New Image() With {.Source = item.Icon},
                        .Tag = item.Tag,
                        .IsEnabled = item.IsEnabled,
                        .FontWeight = item.FontWeight
                    }
                    If Not item.Items Is Nothing Then
                        For Each subItem In getMenuItems(item.Items)
                            menuItem.Items.Add(subItem)
                        Next
                    End If
                    result.Add(menuItem)
                End If
            Next

            Return result
        End Function

        Protected Function getMenuItemData(hMenu2 As IntPtr, parentIndex As Integer) As List(Of MenuItemData)
            Dim result As List(Of MenuItemData) = New List(Of MenuItemData)()

            If Not _contextMenu Is Nothing AndAlso Not IntPtr.Zero.Equals(hMenu2) Then
                If parentIndex >= 0 Then
                    Dim lParam As Integer = (&H0) Or (parentIndex And &HFFFF)
                    If Not _contextMenu3 Is Nothing Then
                        Dim ptr3 As IntPtr, ptr4 As IntPtr
                        Dim h As HRESULT = _contextMenu3.HandleMenuMsg2(WM.INITMENUPOPUP, hMenu2, lParam, ptr3)
                        h = _contextMenu3.HandleMenuMsg2(WM.MENUSELECT, hMenu2, lParam, ptr4)
                    ElseIf Not _contextMenu2 Is Nothing Then
                        _contextMenu2.HandleMenuMsg(WM.INITMENUPOPUP, hMenu2, lParam)
                        _contextMenu2.HandleMenuMsg(WM.MENUSELECT, hMenu2, lParam)
                    End If

                    ' wait for menu to populate
                    Shell.GlobalThreadPool.Run(
                        Sub()
                            Dim initialCount As Integer
                            Do
                                initialCount = Functions.GetMenuItemCount(hMenu2)
                                Thread.Sleep(50)
                            Loop While Functions.GetMenuItemCount(hMenu2) <> initialCount
                        End Sub)
                End If

                For i = 0 To Functions.GetMenuItemCount(hMenu2) - 1
                    Dim mii As MENUITEMINFO
                    mii.cbSize = CUInt(Marshal.SizeOf(mii))
                    mii.fMask = MIIM.MIIM_STRING
                    Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                    Dim header As String
                    mii.cch += 1
                    Try
                        mii.dwTypeData = Marshal.AllocHGlobal(CType(mii.cch * 2, Integer))
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)
                        header = Marshal.PtrToStringUni(mii.dwTypeData)
                    Finally
                        Marshal.FreeHGlobal(mii.dwTypeData)
                    End Try

                    Dim bitmapSource As BitmapSource

                    Try
                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_BITMAP Or MIIM.MIIM_FTYPE Or MIIM.MIIM_CHECKMARKS
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        If Not IntPtr.Zero.Equals(mii.hbmpItem) Then
                            bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpItem, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            bitmapSource.Freeze()
                        ElseIf Not IntPtr.Zero.Equals(mii.hbmpUnchecked) Then
                            bitmapSource = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(mii.hbmpUnchecked, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                            bitmapSource.Freeze()
                        Else
                            bitmapSource = Nothing
                        End If
                    Finally
                        If Not IntPtr.Zero.Equals(mii.hbmpItem) Then
                            If Not _hbitmapsToDispose.Contains(mii.hbmpItem) Then _hbitmapsToDispose.Add(mii.hbmpItem)
                        End If
                        If Not IntPtr.Zero.Equals(mii.hbmpChecked) Then
                            If Not _hbitmapsToDispose.Contains(mii.hbmpChecked) Then _hbitmapsToDispose.Add(mii.hbmpChecked)
                        End If
                        If Not IntPtr.Zero.Equals(mii.hbmpUnchecked) Then
                            If Not _hbitmapsToDispose.Contains(mii.hbmpUnchecked) Then _hbitmapsToDispose.Add(mii.hbmpUnchecked)
                        End If
                    End Try

                    If mii.fType = MFT.SEPARATOR Then
                        ' refuse initial and double separators
                        If Not result.Count = 0 AndAlso Not result(result.Count - 1).Header = "-----" Then
                            result.Add(New MenuItemData() With {.Header = "-----"})
                        End If
                    Else
                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_ID
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim verb As String = Nothing, id As Integer
                        Dim cmd As StringBuilder = New StringBuilder()
                        cmd.Append(New String(" ", 512))
                        If mii.wID >= 1 AndAlso mii.wID <= 99999 Then
                            id = mii.wID - 1
                            _contextMenu.GetCommandString(id, GCS.VERBW, 0, cmd, 512)
                            If cmd.Length = 0 Then
                                _contextMenu.GetCommandString(id, GCS.VERBA, 0, cmd, 512)
                            End If
                            verb = cmd.ToString()
                        End If

                        Debug.WriteLine(header & "  " & id & vbTab & verb?.ToString())

                        mii = New MENUITEMINFO()
                        mii.cbSize = CUInt(Marshal.SizeOf(mii))
                        mii.fMask = MIIM.MIIM_SUBMENU Or MIIM.MIIM_STATE
                        Functions.GetMenuItemInfo(hMenu2, i, True, mii)

                        Dim menuItem As MenuItemData = New MenuItemData() With {
                            .Header = header.Replace("&", "_"),
                            .Icon = bitmapSource,
                            .Tag = New Tuple(Of Integer, String, Object)(id, verb, Nothing),
                            .IsEnabled = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DISABLED), False, True),
                            .FontWeight = If(CType(mii.fState, MFS).HasFlag(MFS.MFS_DEFAULT), FontWeights.Bold, FontWeights.Normal)
                        }

                        If CBool(mii.fState And MFS.MFS_DEFAULT) Then
                            Me.DefaultId = menuItem.Tag
                        End If

                        If Not IntPtr.Zero.Equals(mii.hSubMenu) Then
                            menuItem.Items = getMenuItemData(mii.hSubMenu, i)
                        End If

                        result.Add(menuItem)
                    End If
                Next

                ' remove trailing separators
                While result.Count > 0 AndAlso result(result.Count - 1).Header = "-----"
                    result.RemoveAt(result.Count - 1)
                End While
            End If

            Return result
        End Function

        Protected Overloads Overrides Sub Make(folder As Folder, items As IEnumerable(Of Item), isDefaultOnly As Boolean)
            Dim folderPidl As Pidl = Nothing, itemPidls As Pidl() = Nothing, doUseAbsolutePidls As Boolean

            Shell.GlobalThreadPool.Run(
                Sub()
                    If Not items Is Nothing AndAlso items.Count > 0 Then
                        ' user clicked on an item
                        Dim f As Folder = items(0).Parent
                        If items.All(Function(i) (i.Parent Is Nothing AndAlso f Is Nothing) _
                                         OrElse (Not i.Parent Is Nothing AndAlso If(i.Parent.Pidl?.Equals(f?.Pidl), False))) Then
                            folder = f
                        Else
                            folder = Shell.Desktop
                            doUseAbsolutePidls = True
                        End If
                        If folder Is Nothing Then folder = Shell.Desktop
                        folderPidl = folder.Pidl?.Clone()
                        itemPidls = items.Where(Function(i) Not i.Pidl Is Nothing).Select(Function(i) i.Pidl?.Clone()).ToArray()
                    Else
                        ' user clicked on the background
                        If folder.FullPath = Shell.Desktop.FullPath Then
                            ' this is the desktop
                            folderPidl = Shell.Desktop.Pidl.Clone()
                            itemPidls = {Shell.Desktop.Pidl.Clone()}
                        Else
                            ' this is any other folder
                            If TypeOf folder Is SearchFolder Then
                                folderPidl = Shell.Desktop.Pidl.Clone()
                            Else
                                folderPidl = folder.Parent?.Pidl?.Clone()
                            End If
                            If Not folder.Pidl Is Nothing Then itemPidls = {folder.Pidl?.Clone()}
                        End If
                    End If
                End Sub)

            Dim tcs As TaskCompletionSource = New TaskCompletionSource()
            _thread.Run(
                Sub()
                    Dim flags As CMF = CMF.CMF_NORMAL
                    Dim shellFolder As IShellFolder = Nothing

                    shellFolder = folder.MakeIShellFolderOnCurrentThread()

                    Try
                        If Not shellFolder Is Nothing Then
                            If Not items Is Nothing AndAlso items.Count > 0 Then
                                ' user clicked on an item
                                flags = flags Or CMF.CMF_ITEMMENU

                                CType(shellFolder, IShellFolderForIContextMenu).GetUIObjectOf _
                                    (IntPtr.Zero, itemPidls.Length, itemPidls.Select(Function(p) _
                                        If(doUseAbsolutePidls, p.AbsolutePIDL, p.RelativePIDL)).ToArray(),
                                        GetType(IContextMenu).GUID, 0, _contextMenu)
                            Else
                                ' user clicked on the background
                                CType(shellFolder, IShellFolderForIContextMenu).CreateViewObject _
                                    (IntPtr.Zero, GetType(IContextMenu).GUID, _contextMenu)
                            End If
                        End If

                        If Not _contextMenu Is Nothing Then
                            _contextMenu2 = TryCast(_contextMenu, IContextMenu2)
                            _contextMenu3 = TryCast(_contextMenu, IContextMenu3)
                        End If

                        If Not _contextMenu Is Nothing Then
                            _hMenu = Functions.CreatePopupMenu()
                            flags = flags Or CMF.CMF_EXPLORE Or CMF.CMF_CANRENAME
                            If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then flags = flags Or CMF.CMF_EXTENDEDVERBS
                            If isDefaultOnly Then flags = flags Or CMF.CMF_DEFAULTONLY
                            _contextMenu.QueryContextMenu(_hMenu, 0, 1, 99999, flags)

                            Dim shellExtInit As IShellExtInit, dataObject As IDataObject_PreserveSig = Nothing
                            Try
                                shellExtInit = TryCast(_contextMenu, IShellExtInit)
                                If Not shellExtInit Is Nothing Then
                                    Functions.SHCreateDataObject(folderPidl.AbsolutePIDL, itemPidls.Count,
                                                     itemPidls.Select(Function(p) If(doUseAbsolutePidls, p.AbsolutePIDL, p.RelativePIDL)).ToArray(),
                                                     Nothing, GetType(IDataObject_PreserveSig).GUID, dataObject)
                                    shellExtInit.Initialize(If(folder.Pidl?.AbsolutePIDL, IntPtr.Zero), dataObject, IntPtr.Zero)
                                End If
                            Finally
                                If Not dataObject Is Nothing Then
                                    Marshal.ReleaseComObject(dataObject)
                                    dataObject = Nothing
                                End If
                            End Try
                        End If
                    Finally
                        Shell.GlobalThreadPool.Run(
                            Sub()
                                folderPidl?.Dispose()
                                If Not itemPidls Is Nothing Then
                                    For Each p In itemPidls.Where(Function(p2) Not p2 Is Nothing)
                                        p.Dispose()
                                    Next
                                End If
                            End Sub)
                        If Not shellFolder Is Nothing Then
                            Marshal.ReleaseComObject(shellFolder)
                            shellFolder = Nothing
                        End If
                    End Try
                    tcs.SetResult()
                End Sub)
            tcs.Task.Wait()
        End Sub

        Public Overrides Async Function InvokeCommand(id As Tuple(Of Integer, String, Object)) As Task
            If id Is Nothing Then Return

            Await Make()

            Dim folder As Folder = Nothing
            Dim selectedItems As IEnumerable(Of Item) = Nothing
            Dim e As CommandInvokedEventArgs = Nothing

            If id.Item1.Equals(Me.DefaultId?.Item1) Then
                For Each item In _activeItems
                    If TypeOf item Is Link Then
                        If Not String.IsNullOrWhiteSpace(CType(item, Link).TargetItem.FullPath) Then
                            ' register with os
                            Functions.SHAddToRecentDocs(SHARD.SHARD_PATHW, CType(item, Link).TargetItem.FullPath)
                        End If
                    Else
                        If Not String.IsNullOrWhiteSpace(item.FullPath) Then
                            ' register with os
                            Functions.SHAddToRecentDocs(SHARD.SHARD_PATHW, item.FullPath)
                        End If
                    End If
                Next
            End If

            UIHelper.OnUIThread(
                Sub()
                    folder = Me.Folder
                    selectedItems = Me.SelectedItems

                    e = New CommandInvokedEventArgs() With {
                        .Id = id.Item1,
                        .Verb = id.Item2
                    }
                    Dim c As Control = Me.Buttons.Cast(Of Control).FirstOrDefault(Function(i) id.Equals(i.Tag))
                    If TypeOf c Is ToggleButton Then e.IsChecked = CType(c, ToggleButton).IsChecked
                    RaiseCommandInvoked(e)

                    Select Case id.Item2
                        Case "Windows.ModernShare"
                            Dim __ = Menus.DoShare(selectedItems)
                        Case "copy"
                            Clipboard.CopyFiles(selectedItems)
                        Case "cut"
                            Clipboard.CutFiles(selectedItems)
                    End Select
                End Sub)

            If Not e.IsHandled Then
                _thread.Add(
                    Sub()
                        If id.Item2 = "properties" AndAlso selectedItems.Count > 1 Then
                            id = New Tuple(Of Integer, String, Object)(0, "Laila_Shell_MultiFileProperties", Nothing)
                        End If

                        Select Case id.Item2
                            Case "Windows.ModernShare"
                            Case "Laila_Shell_MultiFileProperties"
                                Dim dataObject As IDataObject_PreserveSig
                                dataObject = Clipboard.GetDataObjectFor(Nothing, selectedItems)
                                Functions.SHMultiFileProperties(dataObject, 0)
                                Marshal.ReleaseComObject(dataObject)
                            Case "copy"
                            Case "cut"
                            Case "paste"
                                If selectedItems Is Nothing OrElse selectedItems.Count = 0 Then
                                    Clipboard.PasteFiles(folder)
                                ElseIf selectedItems?.Count = 1 AndAlso TypeOf selectedItems(0) Is Folder Then
                                    Clipboard.PasteFiles(selectedItems(0))
                                End If
                            Case "delete"
                                Dim dataObject As IDataObject_PreserveSig = Nothing, fo As IFileOperation = Nothing
                                Try
                                    dataObject = Clipboard.GetDataObjectFor(folder, selectedItems.ToList())
                                    fo = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_FileOperation))
                                    If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then fo.SetOperationFlags(FOF.FOFX_WANTNUKEWARNING)
                                    fo.DeleteItems(dataObject)
                                    fo.PerformOperations()
                                Finally
                                    If Not fo Is Nothing Then
                                        Marshal.ReleaseComObject(fo)
                                        fo = Nothing
                                    End If
                                    If Not dataObject Is Nothing Then
                                        Marshal.ReleaseComObject(dataObject)
                                        dataObject = Nothing
                                    End If
                                End Try
                            Case Else
                                folder._wasActivity = True

                                Dim cmi As New CMInvokeCommandInfoEx
                                Debug.WriteLine("InvokeCommand " & id.Item1 & ", " & id.Item2)
                                If id.Item1 > 0 Then
                                    cmi.lpVerb = New IntPtr(id.Item1)
                                    cmi.lpVerbW = New IntPtr(id.Item1)
                                Else
                                    cmi.lpVerb = Marshal.StringToHGlobalAnsi(id.Item2)
                                    cmi.lpVerbW = Marshal.StringToHGlobalUni(id.Item2)
                                End If
                                'cmi.lpDirectory = "c:\temp"
                                'cmi.lpDirectoryW = "c:\temp"
                                cmi.fMask = CMIC.UNICODE Or CMIC.ASYNCOK
                                If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then cmi.fMask = cmi.fMask Or CMIC.CONTROL_DOWN
                                If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then cmi.fMask = cmi.fMask Or CMIC.SHIFT_DOWN
                                cmi.nShow = SW.SHOWNORMAL
                                'cmi.hwnd = Shell._hwnd
                                cmi.cbSize = CUInt(Marshal.SizeOf(cmi))

                                Dim h As HRESULT = _contextMenu.InvokeCommand(cmi)
                                Debug.WriteLine("InvokeCommand returned " & h.ToString())
                        End Select
                    End Sub)
            End If
        End Function

        Protected Overrides Sub Dispose(disposing As Boolean)
            MyBase.Dispose(disposing)

            ' clean up context menu
            If Not _contextMenu Is Nothing Then
                'If Not TypeOf _contextMenu Is IContextMenuImpl Then
                Marshal.ReleaseComObject(_contextMenu)
                'Else
                '    CType(_contextMenu, IContextMenuImpl).Dispose()
                'End If
                _contextMenu = Nothing
            End If

            ' destroy hmenu
            If Not IntPtr.Zero.Equals(_hMenu) Then
                Functions.DestroyMenu(_hMenu)
                _hMenu = IntPtr.Zero
            End If
        End Sub
    End Class
End Namespace