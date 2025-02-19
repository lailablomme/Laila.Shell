Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Media.Imaging
Imports Laila.BalloonTip
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers
Imports Shell32

Namespace Controls
    Public Class Menus
        Inherits FrameworkElement
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly NewItemMenuProperty As DependencyProperty = DependencyProperty.Register("NewItemMenu", GetType(NewItemMenu), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly FolderViewProperty As DependencyProperty = DependencyProperty.Register("FolderView", GetType(FolderView), GetType(Menus), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanCutProperty As DependencyProperty = DependencyProperty.Register("CanCut", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanCopyProperty As DependencyProperty = DependencyProperty.Register("CanCopy", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanPasteProperty As DependencyProperty = DependencyProperty.Register("CanPaste", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanRenameProperty As DependencyProperty = DependencyProperty.Register("CanRename", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanDeleteProperty As DependencyProperty = DependencyProperty.Register("CanDelete", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly CanShareProperty As DependencyProperty = DependencyProperty.Register("CanShare", GetType(Boolean), GetType(Menus), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private disposedValue As Boolean
        Private Shared _rightClickMenu As RightClickMenu

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(Menus), New FrameworkPropertyMetadata(GetType(Menus)))
        End Sub

        Public Sub New()
            AddHandler Shell.ClipboardChanged,
                Async Sub(s As Object, e As EventArgs)
                    Await Me.UpdateButtons()
                End Sub
        End Sub

        Public Shared Async Function InvokeDefaultCommand(item As Item) As Task
            If Not _rightClickMenu Is Nothing Then
                _rightClickMenu.Dispose()
            End If

            _rightClickMenu = New RightClickMenu()
            _rightClickMenu.Folder = If(item.LogicalParent Is Nothing, item, item.LogicalParent)
            _rightClickMenu.SelectedItems = {item}
            _rightClickMenu.IsDefaultOnly = True

            Await _rightClickMenu.Make()
            Await _rightClickMenu.InvokeCommand(_rightClickMenu.DefaultId)
        End Function

        Delegate Sub GetItemNameCoordinatesDelegate(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                         ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)

        Friend Shared Sub DoRename(getCoords As GetItemNameCoordinatesDelegate,
                                   grid As Grid, listBoxItem As ListBoxItem, listBox As ListBox)
            Dim point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double
            Dim item As Item = listBoxItem.DataContext
            Dim originalName As String = Nothing, ext As String = "", isDrive As Boolean, isWithExt As Boolean
            Dim doHideKnownFileExtensions As Boolean = Shell.Settings.DoHideKnownFileExtensions
            Dim balloonTip As BalloonTip.BalloonTip = Nothing
            Dim scrollViewer As ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(listBox)(0)

            ' make sure we get the latest values according to the DoHideKnownFileExtensions setting
            item._fullPath = Nothing
            item._displayName = Nothing

            ' get name and extension
            If item.FullPath.Equals(IO.Path.GetPathRoot(item.FullPath)) Then
                isDrive = True
                SyncLock item._shellItemLock
                    item.ShellItem2.GetDisplayName(SIGDN.PARENTRELATIVEEDITING, originalName)
                End SyncLock
                ext = ""
                isWithExt = True
            ElseIf IO.Path.GetFileName(item.FullPath).StartsWith(item.DisplayName) Then
                originalName = item.DisplayName
                ext = IO.Path.GetFileName(item.FullPath) _
                    .Substring(IO.Path.GetFileNameWithoutExtension(item.FullPath).Length)
                isWithExt = originalName = IO.Path.GetFileName(item.FullPath)
            ElseIf doHideKnownFileExtensions Then
                originalName = IO.Path.GetFileNameWithoutExtension(item.FullPath)
                ext = IO.Path.GetExtension(item.FullPath)
                isWithExt = False
            Else
                originalName = IO.Path.GetFileName(item.FullPath)
                ext = IO.Path.GetExtension(item.FullPath)
                isWithExt = True
            End If

            Dim doRename As Action(Of String) =
                Sub(newName As String)
                    If Not originalName = newName Then
                        ' rename item
                        If isDrive Then
                            Try
                                Dim shellApp As New Shell32.Shell()
                                Dim folder As Shell32.Folder2 = shellApp.NameSpace(item.FullPath)
                                Dim folderItem As FolderItem = folder.Self
                                folderItem.Name = newName
                            Catch ex As Exception
                            End Try
                        Else
                            Dim composedFullName As String = If(isWithExt, newName, newName & ext)
                            Dim fileOperation As IFileOperation = Nothing
                            Try
                                Dim h As HRESULT = Functions.CoCreateInstance(Guids.CLSID_FileOperation, IntPtr.Zero, 1, GetType(IFileOperation).GUID, fileOperation)
                                SyncLock item._shellItemLock
                                    h = fileOperation.RenameItem(item.ShellItem2, composedFullName, Nothing)
                                End SyncLock
                                Debug.WriteLine("RenameItem returned " & h)
                                h = fileOperation.PerformOperations()
                                Debug.WriteLine("PerformOperations returned " & h)
                            Finally
                                If Not fileOperation Is Nothing Then
                                    Marshal.ReleaseComObject(fileOperation)
                                    fileOperation = Nothing
                                End If
                            End Try
                        End If
                    End If
                End Sub

            Dim maxLen As Integer = 260
            If item.IsDrive Then
                Select Case item.PropertiesByCanonicalName("System.Volume.FileSystem").Text
                    Case "NTFS" : maxLen = 32
                    Case "FAT32" : maxLen = 11
                End Select
            End If

            ' get coords
            getCoords(listBoxItem, textAlignment, point, size, fontSize)

            ' make textbox
            Dim textBox As System.Windows.Controls.TextBox
            textBox = New System.Windows.Controls.TextBox() With {
                .Margin = New Thickness(point.X, point.Y, 0, 0),
                .HorizontalAlignment = HorizontalAlignment.Left,
                .VerticalAlignment = VerticalAlignment.Top,
                .Width = size.Width,
                .Height = size.Height,
                .MaxLength = maxLen,
                .TextWrapping = TextWrapping.Wrap,
                .TextAlignment = textAlignment,
                .UseLayoutRounding = True,
                .SnapsToDevicePixels = True,
                .FontSize = fontSize
            }
            textBox.SetValue(Panel.ZIndexProperty, 100)
            textBox.Text = originalName
            grid.Children.Add(textBox)
            textBox.Focus()

            Dim doCancel As Action =
                Sub()
                    ' cancel
                    textBox.Tag = "cancel"
                    grid.Children.Remove(textBox)
                    If Not balloonTip Is Nothing AndAlso balloonTip.IsOpen Then
                        balloonTip.IsOpen = False
                    End If
                End Sub

            ' take scrolling into account
            Dim onScrollChanged As ScrollChangedEventHandler =
                Sub(s As Object, e As ScrollChangedEventArgs)
                    If Not PresentationSource.FromVisual(listBoxItem) Is Nothing Then
                        ' get coords
                        getCoords(listBoxItem, textAlignment, point, size, fontSize)

                        If point.X >= 0 AndAlso point.X + size.Width < listBox.ActualWidth _
                            AndAlso point.Y >= 0 AndAlso point.Y + size.Height < listBox.ActualHeight Then
                            ' we're still visibile, update position
                            textBox.Margin = New Thickness(point.X, point.Y, 0, 0)
                        Else
                            doCancel()
                        End If
                    Else
                        doCancel()
                    End If
                End Sub
            AddHandler scrollViewer.ScrollChanged, onScrollChanged

            ' select filename without extension
            textBox.SelectionStart = 0
            If isWithExt Then
                textBox.SelectionLength = textBox.Text.Length - ext.Length
            Else
                textBox.SelectionLength = textBox.Text.Length
            End If

            ' hook textbox
            AddHandler textBox.LostFocus,
                Sub(s2 As Object, e2 As RoutedEventArgs)
                    RemoveHandler scrollViewer.ScrollChanged, onScrollChanged
                    grid.Children.Remove(textBox)
                    If Not textBox.Tag = "cancel" AndAlso Not String.IsNullOrWhiteSpace(textBox.Text) Then
                        doRename(textBox.Text)
                    End If
                End Sub
            AddHandler textBox.PreviewKeyDown,
                Sub(s2 As Object, e2 As KeyEventArgs)
                    Select Case e2.Key
                        Case Key.Enter
                            grid.Children.Remove(textBox)
                            If Not balloonTip Is Nothing AndAlso balloonTip.IsOpen Then
                                balloonTip.IsOpen = False
                            End If
                        Case Key.Escape
                            doCancel()
                        Case Key.Back
                        Case Else
                            If Keyboard.Modifiers = ModifierKeys.None Then
                                Dim c As Char? = KeyboardHelper.KeyToChar(e2.Key)
                                If Not isDrive AndAlso c.HasValue _
                                AndAlso IO.Path.GetInvalidFileNameChars().Contains(c.Value) Then
                                    If balloonTip Is Nothing Then
                                        balloonTip = New BalloonTip.BalloonTip() With {
                                        .PlacementTarget = textBox,
                                        .Placement = BalloonPlacementMode.Bottom,
                                        .Text = "The following characters can not appear in filenames:" & vbCrLf _
                                              & "     \  /  :  *  ?  ""  <  >  |",
                                        .Timeout = 9000
                                    }
                                        grid.Children.Add(balloonTip)
                                    Else
                                        balloonTip.IsOpen = False
                                    End If
                                    balloonTip.IsOpen = True
                                    e2.Handled = True
                                ElseIf Not balloonTip Is Nothing AndAlso balloonTip.IsOpen Then
                                    balloonTip.IsOpen = False
                                End If
                            End If
                    End Select
                End Sub
        End Sub

        Public Shared Sub DoDelete(items As IEnumerable(Of Item))
            Dim thread As Thread = New Thread(New ThreadStart(
                Sub()
                    Dim fo As IFileOperation = Nothing
                    Dim dataObject As ComTypes.IDataObject = Nothing
                    Try
                        fo = Activator.CreateInstance(Type.GetTypeFromCLSID(Guids.CLSID_FileOperation))
                        dataObject = Clipboard.GetDataObjectFor(items(0).Parent, items)
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
                End Sub))

            thread.SetApartmentState(ApartmentState.STA)
            thread.Start()
        End Sub

        Public Shared Async Function DoShare(items As IEnumerable(Of Item)) As Task
            If (Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDrive) _
                AndAlso items(0).FullPath.StartsWith(Shell.GetSpecialFolder(SpecialFolders.OneDrive).FullPath & IO.Path.DirectorySeparatorChar)) _
                OrElse (Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDriveBusiness) _
                        AndAlso items(0).FullPath.StartsWith(Shell.GetSpecialFolder(SpecialFolders.OneDriveBusiness).FullPath & IO.Path.DirectorySeparatorChar)) Then
                If Not _rightClickMenu Is Nothing Then
                    _rightClickMenu.Dispose()
                End If
                _rightClickMenu = New RightClickMenu() With {
                    .Folder = items(0).Parent,
                    .SelectedItems = items,
                    .IsDefaultOnly = True
                }
                Await _rightClickMenu.Make()
                Await _rightClickMenu.InvokeCommand(New Tuple(Of Integer, String)(0, "{5250E46F-BB09-D602-5891-F476DC89B701}"))
            Else
                Dim assembly As Assembly = Assembly.LoadFrom("Laila.Shell.WinRT.dll")
                Dim type As Type = assembly.GetType("Laila.Shell.WinRT.ModernShare")
                Dim methodInfo As MethodInfo = type.GetMethod("ShowShareUI")
                Dim instance As Object = Activator.CreateInstance(type)
                methodInfo.Invoke(instance, {items.ToList().Select(Function(i) i.FullPath).ToList(),
                                  System.Windows.Application.Current.MainWindow})
            End If
        End Function

        Public Async Function UpdateNewItemMenu() As Task
            If Shell.ShuttingDownToken.IsCancellationRequested Then Return

            If Not Me.NewItemMenu Is Nothing Then
                Me.NewItemMenu.Dispose()
                Me.NewItemMenu = Nothing
            End If

            Dim newItemMenu As NewItemMenu = New NewItemMenu() With {.Folder = Me.Folder}
            Await newItemMenu.Make()
            If newItemMenu.Items.Count > 0 Then
                AddHandler newItemMenu.RenameRequest,
                    Async Sub(s As Object, e As RenameRequestEventArgs)
                        If Not Me.FolderView Is Nothing Then
                            e.IsHandled = Await Me.FolderView.DoRename(e.FullPath)
                        End If
                    End Sub
                Me.NewItemMenu = newItemMenu
            Else
                newItemMenu.Dispose()
                Me.NewItemMenu = Nothing
            End If
        End Function

        Public Async Function UpdateButtons() As Task
            If Not Shell.ShuttingDownToken.IsCancellationRequested Then
                Me.CanCut = Clipboard.CanCut(Me.SelectedItems)
                Me.CanCopy = Clipboard.CanCopy(Me.SelectedItems)
                Me.CanPaste = Not Me.Folder.disposedValue AndAlso Not Me.Folder.IsReadyForDispose AndAlso Not Me.Folder Is Nothing AndAlso Await Clipboard.CanPaste(Me.Folder)
                Me.CanRename = Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 AndAlso Me.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANRENAME))
                Me.CanDelete = Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 AndAlso Me.SelectedItems.All(Function(i) i.Attributes.HasFlag(SFGAO.CANDELETE))
                If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 Then
                    If Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDrive) _
                        AndAlso Me.SelectedItems(0).FullPath.StartsWith(Shell.GetSpecialFolder(SpecialFolders.OneDrive).FullPath & IO.Path.DirectorySeparatorChar) Then
                        Me.CanShare = True
                    ElseIf Shell.GetSpecialFolders().ContainsKey(SpecialFolders.OneDriveBusiness) _
                        AndAlso Me.SelectedItems(0).FullPath.StartsWith(Shell.GetSpecialFolder(SpecialFolders.OneDriveBusiness).FullPath & IO.Path.DirectorySeparatorChar) Then
                        Me.CanShare = True
                    Else
                        Me.CanShare = Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 AndAlso Me.SelectedItems.All(Function(i) IO.File.Exists(i.FullPath))
                    End If
                Else
                    Me.CanShare = False
                End If
            End If
        End Function

        Public Property CanCut As Boolean
            Get
                Return GetValue(CanCutProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanCutProperty, value)
            End Set
        End Property

        Public Property CanCopy As Boolean
            Get
                Return GetValue(CanCopyProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanCopyProperty, value)
            End Set
        End Property

        Public Property CanPaste As Boolean
            Get
                Return GetValue(CanPasteProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanPasteProperty, value)
            End Set
        End Property

        Public Property CanRename As Boolean
            Get
                Return GetValue(CanRenameProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanRenameProperty, value)
            End Set
        End Property

        Public Property CanDelete As Boolean
            Get
                Return GetValue(CanDeleteProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanDeleteProperty, value)
            End Set
        End Property

        Public Property CanShare As Boolean
            Get
                Return GetValue(CanShareProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(CanShareProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property

        Public Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property NewItemMenu As NewItemMenu
            Get
                Return GetValue(NewItemMenuProperty)
            End Get
            Set(value As NewItemMenu)
                SetValue(NewItemMenuProperty, value)
            End Set
        End Property

        Public Property FolderView As FolderView
            Get
                Return GetValue(FolderViewProperty)
            End Get
            Set(value As FolderView)
                SetValue(FolderViewProperty, value)
            End Set
        End Property

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim icm As Menus = TryCast(d, Menus)
            If Not e.NewValue Is Nothing Then
                Await icm.UpdateNewItemMenu()
            End If
            Await icm.UpdateButtons()
        End Sub

        Shared Async Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim icm As Menus = TryCast(d, Menus)
            Await icm.UpdateButtons()
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    ' dispose managed state (managed objects)
                    If Not Me.NewItemMenu Is Nothing Then
                        Me.NewItemMenu.Dispose()
                    End If
                End If

                ' free unmanaged resources (unmanaged objects) and override finalizer
                ' set large fields to null
                disposedValue = True
            End If
        End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub
    End Class
End Namespace