Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Media
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Documents
Imports System.Windows.Input
Imports System.Windows.Media
Imports Laila.Shell.Events
Imports Laila.Shell.Helpers

Namespace Controls
    Public MustInherit Class BaseFolderView
        Inherits Control
        Implements IDisposable

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly ColumnsInProperty As DependencyProperty = DependencyProperty.Register("ColumnsIn", GetType(Behaviors.GridViewExtBehavior.ColumnsInData), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingProperty As DependencyProperty = DependencyProperty.Register("IsLoading", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly IsSelectingProperty As DependencyProperty = DependencyProperty.Register("IsSelecting", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowCheckBoxesToSelectProperty As DependencyProperty = DependencyProperty.Register("DoShowCheckBoxesToSelect", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowCheckBoxesToSelectOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowCheckBoxesToSelectOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowCheckBoxesToSelectOverrideChanged))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared ReadOnly IsDoubleClickToOpenItemProperty As DependencyProperty = DependencyProperty.Register("IsDoubleClickToOpenItem", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDoubleClickToOpenItemOverrideProperty As DependencyProperty = DependencyProperty.Register("IsDoubleClickToOpenItemOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsDoubleClickToOpenItemOverrideChanged))
        Public Shared ReadOnly IsUnderlineItemOnHoverProperty As DependencyProperty = DependencyProperty.Register("IsUnderlineItemOnHover", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsUnderlineItemOnHoverOverrideProperty As DependencyProperty = DependencyProperty.Register("IsUnderlineItemOnHoverOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsUnderlineItemOnHoverOverrideChanged))
        Public Shared ReadOnly DoShowIconsOnlyProperty As DependencyProperty = DependencyProperty.Register("DoShowIconsOnly", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowIconsOnlyOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowIconsOnlyOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowIconsOnlyOverrideChanged))
        Public Shared ReadOnly DoShowTypeOverlayProperty As DependencyProperty = DependencyProperty.Register("DoShowTypeOverlay", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowTypeOverlayOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowTypeOverlayOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowTypeOverlayOverrideChanged))
        Public Shared ReadOnly DoShowFolderContentsInInfoTipProperty As DependencyProperty = DependencyProperty.Register("DoShowFolderContentsInInfoTip", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowFolderContentsInInfoTipOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowFolderContentsInInfoTipOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowFolderContentsInInfoTipOverrideChanged))
        Public Shared ReadOnly DoShowInfoTipsProperty As DependencyProperty = DependencyProperty.Register("DoShowInfoTips", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowInfoTipsOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowInfoTipsOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowInfoTipsOverrideChanged))
        Public Shared ReadOnly IsCompactModeProperty As DependencyProperty = DependencyProperty.Register("IsCompactMode", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsCompactModeOverrideProperty As DependencyProperty = DependencyProperty.Register("IsCompactModeOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsCompactModeOverrideChanged))
        Public Shared ReadOnly DoTypeToSelectProperty As DependencyProperty = DependencyProperty.Register("DoTypeToSelect", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoTypeToSelectOverrideProperty As DependencyProperty = DependencyProperty.Register("DoTypeToSelectOverride", GetType(Boolean?), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoTypeToSelectOverrideChanged))
        Public Shared ReadOnly SearchBoxProperty As DependencyProperty = DependencyProperty.Register("SearchBox", GetType(SearchBox), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly NavigationProperty As DependencyProperty = DependencyProperty.Register("Navigation", GetType(Navigation), GetType(BaseFolderView), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend Host As FolderView
        Friend PART_ListBox As System.Windows.Controls.ListBox
        Private PART_Grid As Grid
        Private PART_CheckBoxSelectAll As CheckBox
        Private _isInternallySettingSelectAll As Boolean
        Private _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _mousePointDown As Point
        Private _mouseItemDown As Item
        Private _mouseItemOver As Item
        Private _mouseOriginalSourceDown As Object
        Private _canOpenWithSingleClick As Boolean
        Protected _scrollViewer As ScrollViewer
        Private _lastScrollOffset As Point
        Private _lastScrollSize As Size
        Private _isLoaded As Boolean
        Private _typeToSearchTimer As Timer
        Private _typeToSearchString As String = ""
        Private _menu As RightClickMenu
        Private _ignoreSelection As Boolean
        Private _toolTip As ToolTip
        Private _toolTipCancellationTokenSource As CancellationTokenSource
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseFolderView), New FrameworkPropertyMetadata(GetType(BaseFolderView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Shell.AddToControlCache(Me)

            Me.PART_ListBox = Template.FindName("PART_ListView", Me)
            Me.PART_Grid = Template.FindName("PART_Grid", Me)
            Me.PART_CheckBoxSelectAll = Template.FindName("PART_CheckBoxSelectAll", Me)
            Dim b = Shell.Settings.DoShowTypeOverlay
            If Not Me.PART_CheckBoxSelectAll Is Nothing Then
                AddHandler Me.PART_CheckBoxSelectAll.Checked,
                    Sub(s As Object, e As RoutedEventArgs)
                        If Not _isInternallySettingSelectAll Then
                            Me.SelectedItems = Me.PART_ListBox.Items.Cast(Of Item)
                            Me.PART_ListBox.Focus()
                        End If
                    End Sub
                AddHandler Me.PART_CheckBoxSelectAll.Unchecked,
                    Sub(s As Object, e As RoutedEventArgs)
                        If Not _isInternallySettingSelectAll Then
                            Me.SelectedItems = Nothing
                            Me.PART_ListBox.Focus()
                        End If
                    End Sub
            End If

            AddHandler Shell.Settings.PropertyChanged,
                Sub(s As Object, e As PropertyChangedEventArgs)
                    Select Case e.PropertyName
                        Case "DoShowCheckBoxesToSelect"
                            setDoShowCheckBoxesToSelect()
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                        Case "IsDoubleClickToOpenItem"
                            setIsDoubleClickToOpenItem()
                        Case "IsUnderlineItemOnHover"
                            setIsUnderlineItemOnHover()
                        Case "DoShowIconsOnly"
                            setDoShowIconsOnly()
                        Case "DoShowTypeOverlay"
                            setDoShowTypeOverlay()
                        Case "DoShowFolderContentsInInfoTip"
                            setDoShowFolderContentsInInfoTip()
                        Case "DoShowInfoTips"
                            setDoShowInfoTips()
                        Case "IsCompactMode"
                            setIsCompactMode()
                        Case "DoTypeToSelect"
                            setDoTypeToSelect()
                    End Select
                End Sub
            setDoShowCheckBoxesToSelect()
            setDoShowEncryptedOrCompressedFilesInColor()
            setIsDoubleClickToOpenItem()
            setIsUnderlineItemOnHover()
            setDoShowIconsOnly()
            setDoShowTypeOverlay()
            setDoShowFolderContentsInInfoTip()
            setDoShowInfoTips()
            setIsCompactMode()
            setDoTypeToSelect()

            AddHandler PART_ListBox.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True
                        Me.PART_ListBox_Loaded()
                    End If
                End Sub

            AddHandler Me.PART_ListBox.PreviewMouseMove, AddressOf OnListViewPreviewMouseMove
            AddHandler Me.PART_ListBox.PreviewMouseDown, AddressOf OnListViewPreviewMouseButtonDown
            AddHandler Me.PART_ListBox.PreviewMouseUp, AddressOf OnListViewPreviewMouseButtonUp
            AddHandler Me.PART_ListBox.MouseLeave, AddressOf OnListViewMouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf OnListViewKeyDown
            AddHandler Me.PreviewTextInput, AddressOf OnListViewTextInput
        End Sub

        Protected Overridable Sub PART_ListBox_Loaded()
            If Not Me.Folder Is Nothing Then
                Me.MakeBinding(Me.Folder)
            End If

            _selectionHelper = New SelectionHelper(Of Item)(Me.PART_ListBox)
            _selectionHelper.SelectionChanged =
                Sub()
                    If Not Me.Folder Is Nothing AndAlso Not _ignoreSelection Then _
                                    Me.SelectedItems = _selectionHelper.SelectedItems
                End Sub
            _selectionHelper.SetSelectedItems(Me.SelectedItems)

            _scrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListBox)(0)
            AddHandler _scrollViewer.ScrollChanged,
                Sub(s2 As Object, e2 As ScrollChangedEventArgs)
                    If Not Me.Folder Is Nothing Then
                        _lastScrollOffset = New Point(_scrollViewer.HorizontalOffset, _scrollViewer.VerticalOffset)
                        _lastScrollSize = New Size(_scrollViewer.ScrollableWidth, _scrollViewer.ScrollableHeight)
                    End If
                End Sub

            If Not Me.Folder Is Nothing Then
                setGrouping(Me.Folder)
            End If
        End Sub

        Public Sub SetSelectedItemsSoft(items As IEnumerable(Of Item))
            _ignoreSelection = True
            _selectionHelper.SetSelectedItems(items)
            _ignoreSelection = False
        End Sub

        Private Sub OnListViewKeyDown(sender As Object, e As KeyEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox AndAlso Not Me.Folder Is Nothing Then
                If e.Key = Key.C AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                    Clipboard.CopyFiles(Me.SelectedItems)
                ElseIf e.Key = Key.X AndAlso Keyboard.Modifiers.HasFlag(ModifierKeys.Control) _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                    Clipboard.CutFiles(Me.SelectedItems)
                ElseIf e.Key = Key.Enter AndAlso Keyboard.Modifiers = ModifierKeys.None _
                AndAlso Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count = 1 Then
                    If TypeOf Me.SelectedItems(0) Is Folder Then
                        Me.Host.Folder = Me.SelectedItems(0)
                    Else
                        invokeDefaultCommand(Me.SelectedItems(0))
                    End If
                    e.Handled = True
                ElseIf e.Key = Key.Back AndAlso Keyboard.Modifiers = ModifierKeys.None Then
                    If Not Me.Navigation Is Nothing AndAlso Me.Navigation.CanBack Then
                        Me.Navigation.Back()
                        e.Handled = True
                    End If
                End If
            End If
        End Sub

        Private Sub OnListViewTextInput(sender As Object, e As TextCompositionEventArgs)
            If Not TypeOf e.OriginalSource Is TextBox AndAlso Not Me.Folder Is Nothing Then
                If Me.DoTypeToSelect Then
                    If Not _typeToSearchTimer Is Nothing Then
                        _typeToSearchTimer.Dispose()
                    End If

                    _typeToSearchTimer = New Timer(New TimerCallback(
                        Sub()
                            UIHelper.OnUIThread(
                            Sub()
                                _typeToSearchString = ""
                                _typeToSearchTimer.Dispose()
                                _typeToSearchTimer = Nothing
                            End Sub)
                        End Sub), Nothing, 650, Timeout.Infinite)

                    _typeToSearchString &= e.Text
                    Dim foundItem As Item
                    If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0 Then
                        foundItem =
                            Me.Folder.Items.Skip(Me.Folder.Items.IndexOf(Me.SelectedItems(0)) + 1) _
                                .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                        If foundItem Is Nothing Then
                            foundItem =
                                Me.Folder.Items.Take(Me.Folder.Items.IndexOf(Me.SelectedItems(0))) _
                                    .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                        End If
                    Else
                        foundItem =
                            Me.Folder.Items _
                                .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                    End If
                    If Not foundItem Is Nothing Then
                        Me.SelectedItems = {foundItem}
                        CType(Me.PART_ListBox.ItemContainerGenerator.ContainerFromItem(foundItem), ListViewItem)?.Focus()
                        e.Handled = True
                    Else
                        SystemSounds.Asterisk.Play()
                    End If
                ElseIf Not Me.SearchBox Is Nothing Then
                    Me.SearchBox.Focus()
                End If
            End If
        End Sub

        Private Sub invokeDefaultCommand(item As Item)
            getMenu(Me.Folder, {item}, True)
            _menu.Make()
            _menu.InvokeCommand(_menu.DefaultId)
        End Sub

        Protected Function GetIsDisplayNameTextBlockTooSmall(textBlock As TextBlock) As Boolean
            textBlock.Measure(New Size(textBlock.ActualWidth,
                If(textBlock.MaxHeight.Equals(Double.NaN), Double.PositiveInfinity, textBlock.MaxHeight)))

            Dim typeface As Typeface = New Typeface(
                textBlock.FontFamily,
                textBlock.FontStyle,
                textBlock.FontWeight,
                textBlock.FontStretch)

            Dim formattedText As FormattedText = New FormattedText(
                textBlock.Inlines.OfType(Of Run)().FirstOrDefault().Text,
                System.Threading.Thread.CurrentThread.CurrentCulture,
                textBlock.FlowDirection,
                typeface,
                textBlock.FontSize,
                textBlock.Foreground,
                VisualTreeHelper.GetDpi(textBlock).PixelsPerDip)
            formattedText.MaxTextWidth = textBlock.ActualWidth
            formattedText.TextAlignment = textBlock.TextAlignment
            formattedText.Trimming = TextTrimming.None

            Return formattedText.Height > textBlock.DesiredSize.Height
        End Function

        Private Sub OnListViewPreviewMouseMove(sender As Object, e As MouseEventArgs)
            Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
            Dim overItem As Item = TryCast(listBoxItem?.DataContext, Item)
            If Me.DoShowInfoTips AndAlso Not overItem Is Nothing AndAlso Not overItem.Equals(_mouseItemOver) Then
                _mouseItemOver = overItem

                If Not _toolTip Is Nothing Then
                    _toolTip.IsOpen = False
                    _toolTip = Nothing
                End If

                If Not _toolTipCancellationTokenSource Is Nothing Then
                    _toolTipCancellationTokenSource.Cancel()
                End If
                _toolTipCancellationTokenSource = New CancellationTokenSource()

                Dim f As Func(Of Task) =
                    Async Function() As Task
                        Dim startTime As DateTime = DateTime.Now

                        Dim startOverItem As Item = overItem
                        Dim text As String = overItem.InfoTip

                        Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                            .FirstOrDefault(Function(b) b.Name = "PART_DisplayName" OrElse b.Tag = "PART_DisplayName")
                        If Not textBlock Is Nothing Then
                            If Me.GetIsDisplayNameTextBlockTooSmall(textBlock) Then
                                text = overItem.DisplayName & Environment.NewLine & text
                            End If
                        End If

                        If Me.DoShowFolderContentsInInfoTip Then
                            Dim textFolderSize As String
                            If TypeOf overItem Is Folder Then
                                textFolderSize = Await CType(overItem, Folder).GetInfoTipFolderSizeAsync(_toolTipCancellationTokenSource.Token)
                            End If
                            If Not String.IsNullOrWhiteSpace(textFolderSize) Then
                                text &= Environment.NewLine & textFolderSize
                            End If
                        End If

                        Await Task.Delay(Math.Max(0, 1500 - DateTime.Now.Subtract(startTime).TotalMilliseconds))

                        If Not String.IsNullOrWhiteSpace(text) AndAlso startOverItem.Equals(_mouseItemOver) Then
                            If Not _toolTip Is Nothing Then
                                _toolTip.IsOpen = False
                                _toolTip = Nothing
                            End If

                            _toolTip = New ToolTip With {
                                .Content = text,
                                .Placement = PlacementMode.Mouse
                            }
                            _toolTip.IsOpen = True
                        End If
                    End Function
                f()
            ElseIf overItem Is Nothing Then
                _mouseItemOver = Nothing
                If Not _toolTip Is Nothing Then
                    _toolTip.IsOpen = False
                    _toolTip = Nothing
                End If
                If Not _toolTipCancellationTokenSource Is Nothing Then
                    _toolTipCancellationTokenSource.Cancel()
                    _toolTipCancellationTokenSource = New CancellationTokenSource()
                End If
            End If

            If Not _mouseItemDown Is Nothing _
                AndAlso (Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Count > 0) _
                AndAlso (e.LeftButton = MouseButtonState.Pressed OrElse e.RightButton = MouseButtonState.Pressed) Then
                Dim currentPointDown As Point = e.GetPosition(Me)
                If Math.Abs(currentPointDown.X - _mousePointDown.X) > 10 OrElse Math.Abs(currentPointDown.Y - _mousePointDown.Y) > 10 Then
                    If Me.SelectedItems Is Nothing OrElse Not Me.SelectedItems.Contains(_mouseItemDown) Then
                        Me.SelectedItems = {_mouseItemDown}
                    End If
                    Drag.Start(Me.SelectedItems, If(e.LeftButton = MouseButtonState.Pressed, MK.MK_LBUTTON, MK.MK_RBUTTON))
                End If
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me)
            _mouseOriginalSourceDown = e.OriginalSource
            _canOpenWithSingleClick = False

            ' this prevents a multiple selection getting replaced by the single clicked item
            If Not e.OriginalSource Is Nothing Then
                Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                _mouseItemDown = clickedItem
                If clickedItem Is Nothing Then
                    Me.PART_ListBox.Focus()
                Else
                    listBoxItem.Focus()
                End If
                If e.LeftButton = MouseButtonState.Pressed AndAlso e.ClickCount = 2 _
                    AndAlso Not clickedItem Is Nothing Then
                    If Me.IsDoubleClickToOpenItem Then
                        Using Shell.OverrideCursor(Cursors.Wait)
                            Me.SelectedItems = {clickedItem}
                            If TypeOf clickedItem Is Folder Then
                                CType(clickedItem, Folder).LastScrollOffset = New Point()
                                Me.Host.Folder = clickedItem
                                UIHelper.OnUIThread(
                            Sub()
                            End Sub, Threading.DispatcherPriority.Render)
                            Else
                                invokeDefaultCommand(clickedItem)
                            End If
                        End Using
                    End If
                ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                    Dim checkBox As CheckBox = UIHelper.GetParentOfType(Of CheckBox)(e.OriginalSource)
                    If Not checkBox Is Nothing Then
                        checkBox.IsChecked = Not checkBox.IsChecked
                        e.Handled = True
                    ElseIf Keyboard.Modifiers = ModifierKeys.None Then
                        If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Contains(clickedItem) Then e.Handled = True
                        _canOpenWithSingleClick = True
                    End If
                ElseIf e.RightButton = MouseButtonState.Pressed AndAlso
                        UIHelper.GetParentOfType(Of Primitives.ScrollBar)(e.OriginalSource) Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of GridViewHeaderRowPresenter)(e.OriginalSource) Is Nothing Then

                    If (Me.SelectedItems Is Nothing OrElse Me.SelectedItems.Count = 0 OrElse Not Me.SelectedItems.Contains(clickedItem)) _
                            AndAlso Not clickedItem Is Nothing Then
                        Me.SelectedItems = {clickedItem}
                    ElseIf clickedItem Is Nothing Then
                        Me.SelectedItems = Nothing
                    End If

                    Me.PART_ListBox.ContextMenu = getMenu(Me.Folder, Me.SelectedItems, False)
                    e.Handled = True
                ElseIf clickedItem Is Nothing AndAlso
                        UIHelper.GetParentOfType(Of System.Windows.Controls.Primitives.ScrollBar)(e.OriginalSource) Is Nothing Then
                    Me.SelectedItems = Nothing
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Public Sub OnListViewPreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
            If _canOpenWithSingleClick _
               AndAlso Not _mouseItemDown Is Nothing AndAlso Not Me.IsDoubleClickToOpenItem Then
                Using Shell.OverrideCursor(Cursors.Wait)
                    Me.SelectedItems = {_mouseItemDown}
                    If TypeOf _mouseItemDown Is Folder Then
                        CType(_mouseItemDown, Folder).LastScrollOffset = New Point()
                        Me.Host.Folder = _mouseItemDown
                        UIHelper.OnUIThread(
                            Sub()
                            End Sub, Threading.DispatcherPriority.Render)
                    Else
                        invokeDefaultCommand(_mouseItemDown)
                    End If
                End Using
            End If

            _mouseItemDown = Nothing
            _canOpenWithSingleClick = False
        End Sub

        Public Sub OnListViewMouseLeave(sender As Object, e As MouseEventArgs)
            _mouseItemDown = Nothing
            _mouseItemOver = Nothing
            If Not _toolTip Is Nothing Then
                _toolTip.IsOpen = False
                _toolTip = Nothing
            End If
            If Not _toolTipCancellationTokenSource Is Nothing Then
                _toolTipCancellationTokenSource.Cancel()
                _toolTipCancellationTokenSource = New CancellationTokenSource()
            End If
        End Sub

        Private Function getMenu(folder As Folder, selectedItems As IEnumerable(Of Item), isDefaultOnly As Boolean) As RightClickMenu
            If Not _menu Is Nothing Then
                _menu.Dispose()
            End If

            _menu = New RightClickMenu()
            _menu.Folder = folder
            _menu.SelectedItems = selectedItems
            _menu.IsDefaultOnly = isDefaultOnly

            AddHandler _menu.CommandInvoked,
                Sub(s As Object, e2 As CommandInvokedEventArgs)
                    Select Case e2.Verb
                        Case "open"
                            If Not Me.SelectedItems Is Nothing _
                                    AndAlso Me.SelectedItems.Count = 1 _
                                    AndAlso TypeOf Me.SelectedItems(0) Is Folder Then
                                Me.Host.Folder = Me.SelectedItems(0)
                                e2.IsHandled = True
                            End If
                        Case "rename"
                            Me.DoRename(Me.SelectedItems(0))
                            e2.IsHandled = True
                        Case "laila.shell.(un)pin"
                            If e2.IsChecked Then
                                PinnedItems.PinItem(Me.SelectedItems(0).FullPath)
                            Else
                                PinnedItems.UnpinItem(Me.SelectedItems(0).FullPath)
                            End If
                            e2.IsHandled = True
                    End Select
                End Sub
            AddHandler _menu.RenameRequest,
                Async Sub(s As Object, e As RenameRequestEventArgs)
                    e.IsHandled = Await Me.DoRename(e.FullPath)
                End Sub

            Return _menu
        End Function

        Protected MustOverride Sub GetItemNameCoordinates(listBoxItem As ListBoxItem, ByRef textAlignment As TextAlignment,
                                                          ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)

        Public Async Function DoRename(fullPath As String) As Task(Of Boolean)
            If Not Me.Folder Is Nothing Then
                Dim item As Item = (Await Me.Folder.GetItemsAsync()).FirstOrDefault(Function(i) i.FullPath?.Equals(fullPath))
                If Not item Is Nothing Then
                    Me.SelectedItems = {item}
                    DoRename(item)
                    Return True
                End If
            End If
            Return False
        End Function

        Public Sub DoRename(item As Item)
            Me.PART_ListBox.ScrollIntoView(item)
            UIHelper.OnUIThread(
                Sub()
                End Sub, Threading.DispatcherPriority.ContextIdle)
            Dim listBoxItem As ListBoxItem = Me.PART_ListBox.ItemContainerGenerator.ContainerFromItem(item)
            DoRename(listBoxItem)
        End Sub

        Public Sub DoRename(listBoxItem As ListBoxItem)
            Dim point As Point, size As Size, textAlignment As TextAlignment, fontSize As Double
            Me.GetItemNameCoordinates(listBoxItem, textAlignment, point, size, fontSize)
            Menus.DoRename(point, size, textAlignment, fontSize, listBoxItem.DataContext, Me.PART_Grid)
        End Sub

        Public Overridable Property SelectedItems As IEnumerable(Of Item)
            Get
                Return GetValue(SelectedItemsProperty)
            End Get
            Set(value As IEnumerable(Of Item))
                SetCurrentValue(SelectedItemsProperty, value)
            End Set
        End Property

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Public Property IsLoading As Boolean
            Get
                Return GetValue(IsLoadingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsLoadingProperty, value)
            End Set
        End Property

        Public Property IsSelecting As Boolean
            Get
                Return GetValue(IsSelectingProperty)
            End Get
            Set(ByVal value As Boolean)
                SetCurrentValue(IsSelectingProperty, value)
            End Set
        End Property

        Public Property DoShowCheckBoxesToSelect As Boolean
            Get
                Return GetValue(DoShowCheckBoxesToSelectProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowCheckBoxesToSelectProperty, value)
            End Set
        End Property

        Private Sub setDoShowCheckBoxesToSelect()
            If Me.DoShowCheckBoxesToSelectOverride.HasValue Then
                Me.DoShowCheckBoxesToSelect = Me.DoShowCheckBoxesToSelectOverride.Value
            Else
                Me.DoShowCheckBoxesToSelect = Shell.Settings.DoShowCheckBoxesToSelect
            End If
        End Sub

        Public Property DoShowCheckBoxesToSelectOverride As Boolean?
            Get
                Return GetValue(DoShowCheckBoxesToSelectOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowCheckBoxesToSelectOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowCheckBoxesToSelectOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowCheckBoxesToSelect()
        End Sub

        Public Property DoShowEncryptedOrCompressedFilesInColor As Boolean
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorProperty, value)
            End Set
        End Property

        Private Sub setDoShowEncryptedOrCompressedFilesInColor()
            If Me.DoShowEncryptedOrCompressedFilesInColorOverride.HasValue Then
                Me.DoShowEncryptedOrCompressedFilesInColor = Me.DoShowEncryptedOrCompressedFilesInColorOverride.Value
            Else
                Me.DoShowEncryptedOrCompressedFilesInColor = Shell.Settings.DoShowEncryptedOrCompressedFilesInColor
            End If
        End Sub

        Public Property DoShowEncryptedOrCompressedFilesInColorOverride As Boolean?
            Get
                Return GetValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowEncryptedOrCompressedFilesInColorOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowEncryptedOrCompressedFilesInColor()
        End Sub

        Public Property IsDoubleClickToOpenItem As Boolean
            Get
                Return GetValue(IsDoubleClickToOpenItemProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(IsDoubleClickToOpenItemProperty, value)
            End Set
        End Property

        Private Sub setIsDoubleClickToOpenItem()
            If Me.IsDoubleClickToOpenItemOverride.HasValue Then
                Me.IsDoubleClickToOpenItem = Me.IsDoubleClickToOpenItemOverride.Value
            Else
                Me.IsDoubleClickToOpenItem = Shell.Settings.IsDoubleClickToOpenItem
            End If
        End Sub

        Public Property IsDoubleClickToOpenItemOverride As Boolean?
            Get
                Return GetValue(IsDoubleClickToOpenItemOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(IsDoubleClickToOpenItemOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnIsDoubleClickToOpenItemOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setIsDoubleClickToOpenItem()
        End Sub

        Public Property IsUnderlineItemOnHover As Boolean
            Get
                Return GetValue(IsUnderlineItemOnHoverProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(IsUnderlineItemOnHoverProperty, value)
            End Set
        End Property

        Private Sub setIsUnderlineItemOnHover()
            If Me.IsUnderlineItemOnHoverOverride.HasValue Then
                Me.IsUnderlineItemOnHover = Me.IsUnderlineItemOnHoverOverride.Value
            Else
                Me.IsUnderlineItemOnHover = Shell.Settings.IsUnderlineItemOnHover
            End If
        End Sub

        Public Property IsUnderlineItemOnHoverOverride As Boolean?
            Get
                Return GetValue(IsUnderlineItemOnHoverOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(IsUnderlineItemOnHoverOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnIsUnderlineItemOnHoverOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setIsUnderlineItemOnHover()
        End Sub

        Public Property DoShowIconsOnly As Boolean
            Get
                Return GetValue(DoShowIconsOnlyProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowIconsOnlyProperty, value)
            End Set
        End Property

        Private Sub setDoShowIconsOnly()
            If Me.DoShowIconsOnlyOverride.HasValue Then
                Me.DoShowIconsOnly = Me.DoShowIconsOnlyOverride.Value
            Else
                Me.DoShowIconsOnly = Shell.Settings.DoShowIconsOnly
            End If
        End Sub

        Public Property DoShowIconsOnlyOverride As Boolean?
            Get
                Return GetValue(DoShowIconsOnlyOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowIconsOnlyOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowIconsOnlyOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowIconsOnly()
        End Sub

        Public Property DoShowTypeOverlay As Boolean
            Get
                Return GetValue(DoShowTypeOverlayProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowTypeOverlayProperty, value)
            End Set
        End Property

        Private Sub setDoShowTypeOverlay()
            If Me.DoShowTypeOverlayOverride.HasValue Then
                Me.DoShowTypeOverlay = Me.DoShowTypeOverlayOverride.Value
            Else
                Me.DoShowTypeOverlay = Shell.Settings.DoShowTypeOverlay
            End If
        End Sub

        Public Property DoShowTypeOverlayOverride As Boolean?
            Get
                Return GetValue(DoShowTypeOverlayOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowTypeOverlayOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowTypeOverlayOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowTypeOverlay()
        End Sub

        Public Property DoShowFolderContentsInInfoTip As Boolean
            Get
                Return GetValue(DoShowFolderContentsInInfoTipProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowFolderContentsInInfoTipProperty, value)
            End Set
        End Property

        Private Sub setDoShowFolderContentsInInfoTip()
            If Me.DoShowFolderContentsInInfoTipOverride.HasValue Then
                Me.DoShowFolderContentsInInfoTip = Me.DoShowFolderContentsInInfoTipOverride.Value
            Else
                Me.DoShowFolderContentsInInfoTip = Shell.Settings.DoShowFolderContentsInInfoTip
            End If
        End Sub

        Public Property DoShowFolderContentsInInfoTipOverride As Boolean?
            Get
                Return GetValue(DoShowFolderContentsInInfoTipOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowFolderContentsInInfoTipOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowFolderContentsInInfoTipOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowFolderContentsInInfoTip()
        End Sub

        Public Property DoShowInfoTips As Boolean
            Get
                Return GetValue(DoShowInfoTipsProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowInfoTipsProperty, value)
            End Set
        End Property

        Private Sub setDoShowInfoTips()
            If Me.DoShowInfoTipsOverride.HasValue Then
                Me.DoShowInfoTips = Me.DoShowInfoTipsOverride.Value
            Else
                Me.DoShowInfoTips = Shell.Settings.DoShowInfoTips
            End If
        End Sub

        Public Property DoShowInfoTipsOverride As Boolean?
            Get
                Return GetValue(DoShowInfoTipsOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowInfoTipsOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowInfoTipsOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoShowInfoTips()
        End Sub

        Public Property IsCompactMode As Boolean
            Get
                Return GetValue(IsCompactModeProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(IsCompactModeProperty, value)
            End Set
        End Property

        Private Sub setIsCompactMode()
            If Me.IsCompactModeOverride.HasValue Then
                Me.IsCompactMode = Me.IsCompactModeOverride.Value
            Else
                Me.IsCompactMode = Shell.Settings.IsCompactMode
            End If
        End Sub

        Public Property IsCompactModeOverride As Boolean?
            Get
                Return GetValue(IsCompactModeOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(IsCompactModeOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnIsCompactModeOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setIsCompactMode()
        End Sub

        Public Property DoTypeToSelect As Boolean
            Get
                Return GetValue(DoTypeToSelectProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoTypeToSelectProperty, value)
            End Set
        End Property

        Private Sub setDoTypeToSelect()
            If Me.DoTypeToSelectOverride.HasValue Then
                Me.DoTypeToSelect = Me.DoTypeToSelectOverride.Value
            Else
                Me.DoTypeToSelect = Shell.Settings.DoTypeToSelect
            End If
        End Sub

        Public Property DoTypeToSelectOverride As Boolean?
            Get
                Return GetValue(DoTypeToSelectOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoTypeToSelectOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoTypeToSelectOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = d
            bfv.setDoTypeToSelect()
        End Sub

        Public Property SearchBox As SearchBox
            Get
                Return GetValue(SearchBoxProperty)
            End Get
            Set(ByVal value As SearchBox)
                SetCurrentValue(SearchBoxProperty, value)
            End Set
        End Property

        Public Property Navigation As Navigation
            Get
                Return GetValue(NavigationProperty)
            End Get
            Set(ByVal value As Navigation)
                SetCurrentValue(NavigationProperty, value)
            End Set
        End Property

        Protected Overridable Sub ClearBinding()
            If Not Me.PART_ListBox Is Nothing Then
                Me.PART_ListBox.ItemsSource = Nothing
            End If
        End Sub

        Protected Overridable Sub MakeBinding(folder As Folder)
            If Not Me.PART_ListBox Is Nothing Then
                setGrouping(folder)

                Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(folder.Items)
                Me.PART_ListBox.ItemsSource = view
            End If
        End Sub

        Protected Overridable Sub Folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Dim folder As Folder = CType(sender, Folder)

            Select Case e.PropertyName
                Case "IsRefreshingItems"
                Case "ItemsSortPropertyName", "ItemsSortDirection", "ItemsGroupByPropertyName", "View"
                    UIHelper.OnUIThread(
                        Sub()
                            If e.PropertyName = "ItemsGroupByPropertyName" Then
                                setGrouping(folder)
                            End If

                            If Not folder.IsRefreshingItems AndAlso Not Me.Folder Is Nothing AndAlso Not TypeOf folder Is SearchFolder Then
                                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder.FullPath)
                                folderViewState.SortPropertyName = folder.ItemsSortPropertyName
                                folderViewState.SortDirection = folder.ItemsSortDirection
                                folderViewState.GroupByPropertyName = folder.ItemsGroupByPropertyName
                                folderViewState.View = folder.View
                                folderViewState.Persist()
                            End If
                        End Sub)
            End Select
        End Sub

        Public Sub Folder_ExpandAllGroups(sender As Object, e As EventArgs)
            Me.ExpandCollapseAllGroups(True)
        End Sub

        Public Sub Folder_CollapseAllGroups(sender As Object, e As EventArgs)
            Me.ExpandCollapseAllGroups(False)
        End Sub

        Public Sub ExpandCollapseAllGroups(isExpanded As Boolean)
            Dim groups As IEnumerable(Of GroupItem) = UIHelper.FindVisualChildren(Of GroupItem)(PART_ListBox)
            For Each group In groups
                Dim toggleButtons As IEnumerable(Of ToggleButton) = UIHelper.FindVisualChildren(Of ToggleButton)(group)
                If toggleButtons.Count > 0 Then
                    toggleButtons(0).IsChecked = isExpanded
                End If
            Next
        End Sub

        Private Sub folder_Items_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            updateCheckBoxSelectAll()
        End Sub

        Private Sub updateCheckBoxSelectAll()
            If Not Me.PART_CheckBoxSelectAll Is Nothing Then
                _isInternallySettingSelectAll = True
                Me.PART_CheckBoxSelectAll.IsChecked =
                    Me.PART_ListBox.Items.Count = If(Me.SelectedItems Is Nothing, 0, Me.SelectedItems.Count)
                _isInternallySettingSelectAll = False
            End If
        End Sub

        Private Sub setGrouping(folder As Folder)
            If Not Me.PART_ListBox.Resources("groupStyle") Is Nothing Then
                If Not String.IsNullOrWhiteSpace(folder.ItemsGroupByPropertyName) Then
                    If Me.PART_ListBox.GroupStyle.Count = 0 Then
                        Me.PART_ListBox.GroupStyle.Add(Me.PART_ListBox.Resources("groupStyle"))
                    End If
                Else
                    Me.PART_ListBox.GroupStyle.Clear()
                End If
            End If
        End Sub

        Shared Async Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseFolderView = TryCast(d, BaseFolderView)

            ' hide listview so no-one sees us binding to the new folder and restoring the scroll position
            If Not bfv.PART_ListBox Is Nothing Then
                bfv.PART_ListBox.Visibility = Visibility.Hidden
            End If

            If Not e.OldValue Is Nothing Then
                Dim oldValue As Folder = e.OldValue

                ' stop listening for events
                RemoveHandler oldValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged
                RemoveHandler oldValue.ExpandAllGroups, AddressOf bfv.Folder_ExpandAllGroups
                RemoveHandler oldValue.CollapseAllGroups, AddressOf bfv.Folder_CollapseAllGroups
                RemoveHandler oldValue.Items.CollectionChanged, AddressOf bfv.folder_Items_CollectionChanged

                ' record last scroll value for use with the back and forward navigation buttons
                oldValue.LastScrollOffset = bfv._lastScrollOffset
                oldValue.LastScrollSize = bfv._lastScrollSize
                If Not bfv._scrollViewer Is Nothing Then
                    bfv._scrollViewer.ScrollToHorizontalOffset(0)
                    bfv._scrollViewer.ScrollToVerticalOffset(0)
                End If

                ' clear view binding
                bfv.ClearBinding()
            End If

            If Not e.NewValue Is Nothing Then
                Dim newValue As Folder = e.NewValue

                ' track recent/frequent folders (in a task because for some folders this might take a while)
                If Not TypeOf newValue Is SearchFolder Then
                    Dim func As Func(Of Task) =
                        Async Function() As Task
                            FrequentFolders.Track(newValue)
                        End Function
                    Task.Run(func)
                End If

                ' set sorting and grouping
                If Not TypeOf newValue Is SearchFolder Then
                    Dim folderViewState As FolderViewState = FolderViewState.FromViewName(newValue.FullPath)
                    newValue.ItemsSortPropertyName = folderViewState.SortPropertyName
                    newValue.ItemsSortDirection = folderViewState.SortDirection
                    newValue.ItemsGroupByPropertyName = folderViewState.GroupByPropertyName
                End If

                ' get notified of folder property changes
                AddHandler newValue.PropertyChanged, AddressOf bfv.Folder_PropertyChanged
                AddHandler newValue.ExpandAllGroups, AddressOf bfv.Folder_ExpandAllGroups
                AddHandler newValue.CollapseAllGroups, AddressOf bfv.Folder_CollapseAllGroups
                AddHandler newValue.Items.CollectionChanged, AddressOf bfv.folder_Items_CollectionChanged

                ' bind view
                bfv.MakeBinding(e.NewValue)

                ' load items
                If Not TypeOf newValue Is SearchFolder Then
                    Await newValue.GetItemsAsync()
                End If

                ' async because otherwise we're a tad too early
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        If Not TypeOf newValue Is SearchFolder Then
                            Await Task.Delay(50)

                            If Not bfv._scrollViewer Is Nothing Then
                                ' restore folder scroll position
                                bfv._lastScrollOffset = newValue.LastScrollOffset
                                bfv._lastScrollSize = newValue.LastScrollSize
                                bfv._scrollViewer.ScrollToHorizontalOffset(If(bfv._lastScrollSize.Width = 0, 0, bfv._lastScrollOffset.X * bfv._scrollViewer.ScrollableWidth / bfv._lastScrollSize.Width))
                                bfv._scrollViewer.ScrollToVerticalOffset(If(bfv._lastScrollSize.Height = 0, 0, bfv._lastScrollOffset.Y * bfv._scrollViewer.ScrollableHeight / bfv._lastScrollSize.Height))
                            End If
                        End If

                        ' show listview
                        bfv.PART_ListBox.Visibility = Visibility.Visible
                    End Sub, Threading.DispatcherPriority.ContextIdle)
            End If
        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim dlv As BaseFolderView = TryCast(d, BaseFolderView)

            If Not dlv._selectionHelper Is Nothing Then
                dlv._selectionHelper.SetSelectedItems(e.NewValue)
            End If
            dlv.updateCheckBoxSelectAll()
        End Sub

        Protected Overridable Sub Dispose(disposing As Boolean)
            If Not disposedValue Then
                If disposing Then
                    If Not _typeToSearchTimer Is Nothing Then
                        _typeToSearchTimer.Dispose()
                        _typeToSearchTimer = Nothing
                    End If
                    If Not Me.Folder Is Nothing Then
                        Me.Folder.IsActiveInFolderView = False
                        Me.Folder = Nothing
                    End If

                    Shell.RemoveFromControlCache(Me)
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