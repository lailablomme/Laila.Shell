Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Media
Imports System.Security.AccessControl
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
Imports Laila.Shell.Interfaces
Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.Application
Imports Laila.Shell.Interop.Windows
Imports WpfToolkit.Controls

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
        Public Shared ReadOnly ScrollOffsetProperty As DependencyProperty = DependencyProperty.Register("ScrollOffset", GetType(Point), GetType(BaseFolderView), New FrameworkPropertyMetadata(New Point(), FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ExpandCollapseAllStateProperty As DependencyProperty = DependencyProperty.Register("ExpandCollapseAllState", GetType(Boolean), GetType(BaseFolderView), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Friend Host As FolderView
        Private PART_CheckBoxSelectAll As CheckBox
        Protected PART_DragInsertIndicator As Grid
        Protected PART_Grid As Grid
        Friend PART_ListBox As ListBox
        Public PART_ScrollViewer As ScrollViewer
        Private _canOpenWithSingleClick As Boolean
        Private _doCancelScroll As Boolean = False
        Private _dragViewStrategy As IDragViewStrategy
        Private _ignoreSelection As Boolean
        Private _isInternallySettingSelectAll As Boolean
        Private _isKeyboardScrolling As Boolean
        Private _isKeyboardSelecting As Boolean
        Private _isLoaded As Boolean
        Private _keyboardSelectingLock As SemaphoreSlim = New SemaphoreSlim(1, 1)
        Private _lastScrollSize As Size
        Private _menu As BaseMenu
        Private _mouseItemDown As Item
        Private _mouseItemOver As Item
        Private _mouseOriginalSourceDown As Object
        Private _mouseOverTime As DateTime
        Private _mousePointDown As Point
        Protected _selectionHelper As SelectionHelper(Of Item) = Nothing
        Private _toolTip As ToolTip
        Private _toolTipCancellationTokenSource As CancellationTokenSource
        Private _typeToSearchString As String = ""
        Private _typeToSearchTimer As Timer
        Private disposedValue As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseFolderView), New FrameworkPropertyMetadata(GetType(BaseFolderView)))
        End Sub

        Public Sub New()
            EventManager.RegisterClassHandler(GetType(FrameworkElement), FrameworkElement.RequestBringIntoViewEvent,
                                              New RequestBringIntoViewEventHandler(AddressOf OnRequestBringIntoView))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            Shell.AddToControlCache(Me)

            Me.PART_ListBox = Template.FindName("PART_ListView", Me)
            Me.PART_Grid = Template.FindName("PART_Grid", Me)
            Me.PART_CheckBoxSelectAll = Template.FindName("PART_CheckBoxSelectAll", Me)
            Me.PART_DragInsertIndicator = Template.FindName("PART_DragInsertIndicator", Me)

            Dim b = Shell.Settings.DoShowTypeOverlay

            If Not Me.PART_CheckBoxSelectAll Is Nothing Then
                AddHandler Me.PART_CheckBoxSelectAll.Checked,
                    Sub(s As Object, e As RoutedEventArgs)
                        If Not _isInternallySettingSelectAll Then
                            Using Shell.OverrideCursor(Cursors.Wait)
                                Me.SelectedItems = Me.PART_ListBox.Items.Cast(Of Item)
                                Me.PART_ListBox.Focus()
                            End Using
                        End If
                    End Sub
                AddHandler Me.PART_CheckBoxSelectAll.Unchecked,
                    Sub(s As Object, e As RoutedEventArgs)
                        If Not _isInternallySettingSelectAll Then
                            Using Shell.OverrideCursor(Cursors.Wait)
                                Me.SelectedItems = Nothing
                                Me.PART_ListBox.Focus()
                            End Using
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

            AddHandler Me.PreviewGotKeyboardFocus,
                Sub(s As Object, e As KeyboardFocusChangedEventArgs)
                End Sub

            AddHandler Me.PART_ListBox.PreviewMouseMove, AddressOf listBox_PreviewMouseMove
            AddHandler Me.PART_ListBox.PreviewMouseDown, AddressOf listBox_PreviewMouseButtonDown
            AddHandler Me.PART_ListBox.PreviewMouseUp, AddressOf listBox_PreviewMouseButtonUp
            AddHandler Me.PART_ListBox.MouseLeave, AddressOf listBox_MouseLeave
            AddHandler Me.PreviewKeyDown, AddressOf baseFolderView_KeyDown
            AddHandler Me.PreviewTextInput, AddressOf baseFolderView_TextInput
        End Sub

        Protected Overridable Sub PART_ListBox_Loaded()
            If Not Me.Folder Is Nothing Then
                Me.MakeBinding(Me.Folder)
            End If

            _selectionHelper = New SelectionHelper(Of Item)(Me.PART_ListBox)
            _selectionHelper.SelectionChanged =
                Sub()
                    If Not _selectionHelper Is Nothing AndAlso Not Me.Folder Is Nothing AndAlso Not _ignoreSelection Then _
                        Me.SelectedItems = _selectionHelper.SelectedItems
                End Sub
            _selectionHelper.SetSelectedItems(Me.SelectedItems)

            Me.PART_ScrollViewer = UIHelper.FindVisualChildren(Of ScrollViewer)(Me.PART_ListBox)(0)
            AddHandler Me.PART_ScrollViewer.ScrollChanged,
                Sub(s2 As Object, e2 As ScrollChangedEventArgs)
                    _doCancelScroll = True
                    If Not Me.Folder Is Nothing Then
                        Me.ScrollOffset = New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.VerticalOffset)
                        _lastScrollSize = New Size(PART_ScrollViewer.ScrollableWidth, PART_ScrollViewer.ScrollableHeight)
                    End If
                End Sub

            If Not Me.Folder Is Nothing Then
                setGrouping(Me.Folder)
            End If

            If Me.IsKeyboardFocusWithin Then
                Me.PART_ListBox.Focus()
            End If
        End Sub

        Public Function GetListBoxClientSize() As Size
            If Not PART_ScrollViewer Is Nothing Then
                Dim widthWithoutScroll As Double = Me.PART_ListBox.ActualWidth
                Dim heightWithoutScroll As Double = Me.PART_ListBox.ActualHeight

                Dim vs As ScrollBar = PART_ScrollViewer.Template.FindName("PART_VerticalScrollBar", PART_ScrollViewer)
                Dim hs As ScrollBar = PART_ScrollViewer.Template.FindName("PART_HorizontalScrollBar", PART_ScrollViewer)

                ' Check visibility of vertical scrollbar
                If vs IsNot Nothing AndAlso vs.Visibility = Visibility.Visible Then
                    widthWithoutScroll -= vs.ActualWidth
                End If

                ' Check visibility of horizontal scrollbar
                If hs IsNot Nothing AndAlso hs.Visibility = Visibility.Visible Then
                    heightWithoutScroll -= hs.ActualHeight
                End If

                Return New Size(Math.Max(widthWithoutScroll, 0), Math.Max(heightWithoutScroll, 0))
            End If

            ' Return the full size if no ScrollViewer was found
            Return New Size(Me.PART_ListBox.ActualWidth, Me.PART_ListBox.ActualHeight)
        End Function

        Public Overridable Sub SetSelectedItemsSoft(items As IEnumerable(Of Item))
            _ignoreSelection = True
            _selectionHelper.SetSelectedItems(items, False)
            _ignoreSelection = False
        End Sub

        Private Sub baseFolderView_KeyDown(sender As Object, e As KeyEventArgs)
            Debug.WriteLine(e.Key)
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
                        Dim __ = invokeDefaultCommand(Me.SelectedItems(0))
                    End If
                    e.Handled = True
                ElseIf e.Key = Key.Back AndAlso Keyboard.Modifiers = ModifierKeys.None Then
                    If Not Me.Navigation Is Nothing AndAlso Me.Navigation.CanBack Then
                        Me.Navigation.Back()
                        e.Handled = True
                    End If
                ElseIf e.Key = Key.Up Then
                    focusAdjacentItem(0, -10)
                    e.Handled = True
                ElseIf e.Key = Key.Down Then
                    focusAdjacentItem(0, +10)
                    e.Handled = True
                ElseIf e.Key = Key.Left Then
                    e.Handled = focusAdjacentItem(-10, 0)
                ElseIf e.Key = Key.Right Then
                    e.Handled = focusAdjacentItem(10, 0)
                End If
            End If
        End Sub

        Private Function focusAdjacentItem(offsetX As Integer, offsetY As Integer) As Boolean
            If _isKeyboardScrolling Then Return False

            If TypeOf Keyboard.FocusedElement Is ListBoxItem OrElse TypeOf Keyboard.FocusedElement Is Button Then
                Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(Folder.Items)

                Dim isFirstItem As Func(Of FrameworkElement, Boolean) =
                    Function(item As FrameworkElement) As Boolean
                        If TypeOf item Is ListBoxItem Then
                            Return item.DataContext.Equals(view(0))
                        Else
                            Dim exp As SlidingExpander = UIHelper.GetParentOfType(Of SlidingExpander)(item)
                            Dim group As CollectionViewGroup = exp.DataContext
                            Return group.Items.Contains(view(0))
                        End If
                    End Function
                Dim isLastItem As Func(Of FrameworkElement, Boolean) =
                    Function(item As FrameworkElement) As Boolean
                        If TypeOf item Is ListBoxItem Then
                            Return item.DataContext.Equals(view(view.Count - 1))
                        Else
                            Dim exp As SlidingExpander = UIHelper.GetParentOfType(Of SlidingExpander)(item)
                            Dim group As CollectionViewGroup = exp.DataContext
                            Return group.Items.Contains(view(view.Count - 1))
                        End If
                    End Function

                ' get listboxitem
                Dim currentItem As FrameworkElement = Keyboard.FocusedElement
                currentItem.BringIntoView(New Rect(0, 0, currentItem.ActualWidth, currentItem.ActualHeight))

                ' get orientation
                Dim orientation As Orientation = Orientation.Vertical
                Dim wrapPanel As WrapPanel = UIHelper.GetParentOfType(Of WrapPanel)(currentItem)
                If Not wrapPanel Is Nothing Then
                    orientation = wrapPanel.Orientation
                Else
                    Dim virtualizingWrapPanel As VirtualizingWrapPanel = UIHelper.GetParentOfType(Of VirtualizingWrapPanel)(currentItem)
                    If Not virtualizingWrapPanel Is Nothing Then
                        orientation = virtualizingWrapPanel.Orientation
                    Else
                        Dim stackPanel As StackPanel = UIHelper.GetParentOfType(Of StackPanel)(currentItem)
                        If Not stackPanel Is Nothing Then
                            orientation = stackPanel.Orientation
                        Else
                            Dim virtualizingStackPanel As VirtualizingStackPanel = UIHelper.GetParentOfType(Of VirtualizingStackPanel)(currentItem)
                            If Not virtualizingStackPanel Is Nothing Then
                                orientation = virtualizingStackPanel.Orientation
                            End If
                        End If
                    End If
                End If

                ' get headerrowpresenter
                Dim headerRowPresenter As GridViewHeaderRowPresenter = UIHelper.FindVisualChildren(Of GridViewHeaderRowPresenter)(Me.PART_ListBox)(0)

                ' init
                Dim ptCurrentItem As Point = currentItem.TranslatePoint(New Point(0, 0), Me.PART_ListBox)
                Dim vp As VirtualizingPanel = UIHelper.GetParentOfType(Of VirtualizingPanel)(currentItem)
                Dim doContinue As Boolean = True
                Dim adjacentItem As FrameworkElement = Nothing
                Dim pt As Point = New Point(ptCurrentItem.X + offsetX + If(offsetX = 0, 5, If(offsetX <= 0, 0, currentItem.ActualWidth)),
                                            ptCurrentItem.Y + offsetY + If(offsetY = 0, currentItem.ActualHeight / 2, If(offsetY <= 0, 0, currentItem.ActualHeight)))

                ' find adjacent item
                Dim doFindOnDifferentX As Boolean = False, doFindOnDifferentY As Boolean = False
                Do
                    'pt.X = Math.Min(Math.Max(2, pt.X), _scrollViewer.ViewportWidth - 2)
                    'pt.Y = Math.Min(Math.Max(2 + If(headerRowPresenter?.ActualHeight, 0), pt.Y), _scrollViewer.ViewportHeight - 2 + If(headerRowPresenter?.ActualHeight, 0))
                    For scanCount = 0 To If(offsetX <> 0, currentItem.ActualHeight, currentItem.ActualWidth) / 10
                        Dim ptScan As Point = New Point(pt.X + If(offsetY > 0, 5, If(offsetY < 0, -5, 0)) * scanCount, pt.Y + If(offsetX > 0, 5, If(offsetX < 0, -5, 0)) * scanCount)
                        Dim el As IInputElement = Me.PART_ListBox.InputHitTest(ptScan)
                        If Not el Is Nothing Then
                            adjacentItem = UIHelper.GetParentOfType(Of ListBoxItem)(el)
                            If adjacentItem Is Nothing AndAlso (offsetX = 0 OrElse (TypeOf currentItem Is Button AndAlso orientation = Orientation.Horizontal)) Then
                                adjacentItem = UIHelper.GetParentOfType(Of SlidingExpander)(el)
                                If Not adjacentItem Is Nothing Then
                                    Dim transform As GeneralTransform = Me.PART_ListBox.TransformToDescendant(adjacentItem)
                                    Dim itemPoint As Point = transform.Transform(ptScan)
                                    If itemPoint.X > 30 AndAlso orientation = Orientation.Vertical Then
                                        adjacentItem = Nothing
                                    Else
                                        Dim button As Button = UIHelper.FindVisualChildren(Of Button)(adjacentItem)?(0)
                                        If Not button Is Nothing Then
                                            If itemPoint.Y > button.ActualHeight Then
                                                adjacentItem = Nothing
                                            End If
                                        End If
                                    End If
                                    If Not adjacentItem Is Nothing Then
                                        adjacentItem = UIHelper.FindVisualChildren(Of Button)(adjacentItem)?(0)
                                    End If
                                End If
                            End If
                        End If
                        If currentItem.Equals(adjacentItem) Then adjacentItem = Nothing
                        If Not UIHelper.IsAncestor(Me.PART_ListBox, adjacentItem) Then adjacentItem = Nothing
                        If Not adjacentItem Is Nothing Then Exit For
                    Next
                    Dim ptAdjacentItem As Point = New Point(0, 0)
                    If Not adjacentItem Is Nothing Then
                        ptAdjacentItem = adjacentItem.TranslatePoint(New Point(0, 0), Me.PART_ListBox)
                    End If
                    If doFindOnDifferentX AndAlso ptAdjacentItem.X = ptCurrentItem.X Then
                        If pt.X < ptCurrentItem.X Then pt.X -= 10 Else pt.X += 10
                        If pt.X <= 2 Then scrollTo(New Point(PART_ScrollViewer.HorizontalOffset - 10, PART_ScrollViewer.VerticalOffset))
                        If pt.X >= PART_ScrollViewer.ViewportWidth - 2 Then scrollTo(New Point(PART_ScrollViewer.HorizontalOffset + 10, PART_ScrollViewer.VerticalOffset))
                    ElseIf doFindOnDifferentY AndAlso ptAdjacentItem.Y = ptCurrentItem.Y Then
                        If pt.Y < ptCurrentItem.Y Then pt.Y -= 10 Else pt.Y += 10
                        If pt.Y <= 2 Then scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.VerticalOffset - 10))
                        If pt.Y >= PART_ScrollViewer.ViewportHeight - 2 + If(headerRowPresenter?.ActualHeight, 0) Then
                            scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.VerticalOffset + 10))
                        End If
                    ElseIf Not adjacentItem Is Nothing AndAlso Not adjacentItem.Equals(currentItem) Then
                        doContinue = False
                    ElseIf offsetX <> 0 AndAlso pt.X > 2 AndAlso pt.X < PART_ScrollViewer.ViewportWidth - 2 Then
                        pt.X += offsetX
                    ElseIf offsetY <> 0 AndAlso pt.Y > 2 + If(headerRowPresenter?.ActualHeight, 0) AndAlso pt.Y < PART_ScrollViewer.ViewportHeight - 2 + If(headerRowPresenter?.ActualHeight, 0) Then
                        pt.Y += offsetY
                    ElseIf offsetX < 0 AndAlso PART_ScrollViewer.HorizontalOffset > 0 Then
                        scrollTo(New Point(PART_ScrollViewer.HorizontalOffset - 250, PART_ScrollViewer.VerticalOffset))
                    ElseIf offsetY < 0 AndAlso PART_ScrollViewer.VerticalOffset > 0 AndAlso Not Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled Then
                        scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.VerticalOffset - 10))
                    ElseIf offsetX > 0 AndAlso PART_ScrollViewer.HorizontalOffset < PART_ScrollViewer.ScrollableWidth Then
                        scrollTo(New Point(PART_ScrollViewer.HorizontalOffset + 10, PART_ScrollViewer.VerticalOffset))
                    ElseIf offsetY > 0 AndAlso PART_ScrollViewer.VerticalOffset < PART_ScrollViewer.ScrollableHeight AndAlso Not Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled Then
                        scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.VerticalOffset + 10))
                    ElseIf offsetX < 0 AndAlso orientation = Orientation.Horizontal Then
                        If pt.Y - currentItem.ActualHeight < 0 Then
                            If PART_ScrollViewer.VerticalOffset - 10 > 0 AndAlso PART_ScrollViewer.ViewportHeight > currentItem.ActualHeight Then
                                scrollTo(New Point(PART_ScrollViewer.ScrollableWidth, PART_ScrollViewer.VerticalOffset - 10))
                                pt.Y -= currentItem.ActualHeight
                                pt.X = If(Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled, 0, PART_ScrollViewer.ViewportWidth - 3)
                            Else
                                ' we're out of options
                                doContinue = False
                            End If
                        ElseIf Not isFirstItem(currentItem) Then
                            scrollTo(New Point(PART_ScrollViewer.ScrollableWidth, PART_ScrollViewer.VerticalOffset))
                            pt.Y -= currentItem.ActualHeight
                            pt.X = If(Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled, 0, PART_ScrollViewer.ViewportWidth - 3)
                        Else
                            ' we're out of options
                            doContinue = False
                        End If
                        doFindOnDifferentY = True
                    ElseIf offsetY < 0 Then
                        If pt.X - 35 < 0 Then
                            If PART_ScrollViewer.HorizontalOffset - 35 > 0 AndAlso PART_ScrollViewer.ViewportWidth > currentItem.ActualWidth Then
                                scrollTo(New Point(PART_ScrollViewer.HorizontalOffset - 35, PART_ScrollViewer.ScrollableHeight))
                                pt.X -= 35
                                pt.Y = PART_ScrollViewer.ViewportHeight - 3
                            Else
                                ' we're out of options
                                doContinue = False
                            End If
                        ElseIf Not isFirstItem(currentItem) Then
                            scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, PART_ScrollViewer.ScrollableHeight))
                            pt.X -= 35
                            pt.Y = PART_ScrollViewer.ViewportHeight - 3
                        Else
                            ' we're out of options
                            doContinue = False
                        End If
                        doFindOnDifferentX = True
                    ElseIf offsetX > 0 AndAlso orientation = Orientation.Horizontal Then
                        If pt.Y + currentItem.ActualHeight > PART_ScrollViewer.ViewportHeight AndAlso Not Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled Then
                            If PART_ScrollViewer.VerticalOffset + 10 < PART_ScrollViewer.ScrollableHeight AndAlso PART_ScrollViewer.ViewportHeight > currentItem.ActualHeight _
                                AndAlso Not Me.PART_ScrollViewer.VerticalScrollBarVisibility = ScrollBarVisibility.Disabled Then
                                scrollTo(New Point(0, PART_ScrollViewer.VerticalOffset + 10))
                                pt.Y += currentItem.ActualHeight
                                pt.X = 3
                            Else
                                ' we're out of options
                                doContinue = False
                            End If
                        ElseIf Not isLastItem(currentItem) Then
                            scrollTo(New Point(0, PART_ScrollViewer.VerticalOffset))
                            pt.Y += currentItem.ActualHeight
                            pt.X = 3
                        Else
                            ' we're out of options
                            doContinue = False
                        End If
                        doFindOnDifferentY = True
                    ElseIf offsetY > 0 AndAlso orientation = Orientation.Vertical Then
                        If pt.X + currentItem.ActualWidth > PART_ScrollViewer.ViewportWidth Then
                            If PART_ScrollViewer.HorizontalOffset + 25 < PART_ScrollViewer.ScrollableWidth AndAlso PART_ScrollViewer.ViewportWidth > currentItem.ActualWidth Then
                                scrollTo(New Point(PART_ScrollViewer.HorizontalOffset + 25, 0))
                                pt.X += currentItem.ActualWidth
                                pt.Y = 3
                            Else
                                ' we're out of options
                                doContinue = False
                            End If
                        ElseIf Not isLastItem(currentItem) Then
                            scrollTo(New Point(PART_ScrollViewer.HorizontalOffset, 0))
                            pt.X += currentItem.ActualWidth
                            pt.Y = 3
                        Else
                            ' we're out of options
                            doContinue = False
                        End If
                        doFindOnDifferentX = True
                    Else
                        ' we're out of options
                        doContinue = False
                    End If
                    pt.X = Math.Min(Math.Max(2, pt.X), PART_ScrollViewer.ViewportWidth - 2)
                    pt.Y = Math.Min(Math.Max(2 + If(headerRowPresenter?.ActualHeight, 0), pt.Y), PART_ScrollViewer.ViewportHeight - 2 + If(headerRowPresenter?.ActualHeight, 0))
                Loop While doContinue

                ' select
                If Not adjacentItem Is Nothing Then
                    If TypeOf adjacentItem Is ListBoxItem Then
                        adjacentItem.Focus()
                        If Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then
                            adjacentItem.BringIntoView(New Rect(0, 0, adjacentItem.ActualWidth, adjacentItem.ActualHeight))
                        End If
                    Else
                        Dim parentItem As SlidingExpander = UIHelper.GetParentOfType(Of SlidingExpander)(adjacentItem)
                        If Not parentItem Is Nothing Then
                            parentItem.Focus()
                            parentItem.BringIntoView(New Rect(0, 0, adjacentItem.ActualWidth, adjacentItem.ActualHeight))
                        End If
                    End If
                    If Keyboard.Modifiers.HasFlag(ModifierKeys.Shift) Then
                    ElseIf Not Keyboard.Modifiers.HasFlag(ModifierKeys.Control) Then
                        If TypeOf adjacentItem Is ListBoxItem Then Me.SelectedItems = {adjacentItem.DataContext}
                    End If
                    Return Not adjacentItem.Equals(currentItem)
                End If
            End If

            Return False
        End Function

        Protected Sub scrollTo(pt As Point)
            _isKeyboardScrolling = True
            pt.X = Math.Min(Math.Max(0, pt.X), PART_ScrollViewer.ScrollableWidth)
            pt.Y = Math.Min(Math.Max(0, pt.Y), PART_ScrollViewer.ScrollableHeight)
            UIHelper.OnUIThread(
                Sub()
                    _doCancelScroll = False
                    PART_ScrollViewer.ScrollToHorizontalOffset(pt.X)
                    PART_ScrollViewer.ScrollToVerticalOffset(pt.Y)
                End Sub, Threading.DispatcherPriority.Input)
            'Do While Not _doCancelScroll AndAlso (Math.Abs(_scrollViewer.HorizontalOffset - pt.X) > 0.5 OrElse Math.Abs(_scrollViewer.VerticalOffset - pt.Y) > 0.5)
            UIHelper.OnUIThread(
                Sub()
                End Sub, Threading.DispatcherPriority.Render)
            'Loop
            _isKeyboardScrolling = False
        End Sub

        Protected Overridable Sub OnRequestBringIntoView(s As Object, e As RequestBringIntoViewEventArgs)
            If TypeOf e.OriginalSource Is ListBoxItem AndAlso UIHelper.IsAncestor(Me.PART_ListBox, e.OriginalSource) Then
                Dim item As ListBoxItem = e.OriginalSource
                If Not item Is Nothing AndAlso UIHelper.IsAncestor(PART_ScrollViewer, item) Then
                    Dim transform As GeneralTransform = item.TransformToAncestor(PART_ScrollViewer)
                    Dim itemRect As Rect = transform.TransformBounds(New Rect(0, 0, item.ActualWidth, item.ActualHeight))

                    ' Check if item is outside the viewport and adjust scrolling 
                    If itemRect.Top < 0 Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + itemRect.Top)
                    ElseIf itemRect.Bottom > PART_ScrollViewer.ViewportHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + (itemRect.Bottom - PART_ScrollViewer.ViewportHeight))
                    End If
                    If itemRect.Left < 0 Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + itemRect.Left)
                    ElseIf itemRect.Right > PART_ScrollViewer.ViewportWidth Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + (itemRect.Right - PART_ScrollViewer.ViewportWidth))
                    End If
                    e.Handled = True
                End If
            ElseIf TypeOf e.OriginalSource Is SlidingExpander AndAlso UIHelper.IsAncestor(Me.PART_ListBox, e.OriginalSource) Then
                Dim parentItem As SlidingExpander = e.OriginalSource
                Dim buttonItem As Button = UIHelper.FindVisualChildren(Of Button)(parentItem)?(0)
                Dim item As Border = UIHelper.GetParentOfType(Of Border)(buttonItem)
                If Not item Is Nothing AndAlso UIHelper.IsAncestor(PART_ScrollViewer, item) Then
                    Dim transform As GeneralTransform = item.TransformToAncestor(PART_ScrollViewer)
                    Dim itemRect As Rect = transform.TransformBounds(New Rect(0, 0, item.ActualWidth, item.ActualHeight))
                    ' Check if item is outside the viewport and adjust scrolling 
                    If itemRect.Top < 0 Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + itemRect.Top)
                    ElseIf itemRect.Bottom > PART_ScrollViewer.ViewportHeight Then
                        PART_ScrollViewer.ScrollToVerticalOffset(PART_ScrollViewer.VerticalOffset + (itemRect.Bottom - PART_ScrollViewer.ViewportHeight))
                    End If
                    If itemRect.Left < 0 Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + itemRect.Left)
                    ElseIf itemRect.Right > PART_ScrollViewer.ViewportWidth Then
                        PART_ScrollViewer.ScrollToHorizontalOffset(PART_ScrollViewer.HorizontalOffset + (itemRect.Right - PART_ScrollViewer.ViewportWidth))
                    End If
                    e.Handled = True
                End If
            ElseIf Not If(TypeOf e.OriginalSource Is Expander, e.OriginalSource, UIHelper.GetParentOfType(Of Expander)(e.OriginalSource)) Is Nothing _
                AndAlso UIHelper.GetParentOfType(Of ListBox)(e.OriginalSource)?.Equals(Me.PART_ListBox) Then
                e.Handled = True
            End If
        End Sub

        Private Sub baseFolderView_TextInput(sender As Object, e As TextCompositionEventArgs)
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
                            Me.Folder.Items.ToList().Skip(Me.Folder.Items.ToList().IndexOf(Me.SelectedItems(0)) + 1) _
                                .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                        If foundItem Is Nothing Then
                            foundItem =
                                Me.Folder.Items.ToList().Take(Me.Folder.Items.ToList().IndexOf(Me.SelectedItems(0))) _
                                    .FirstOrDefault(Function(i) i.DisplayName.ToLower().StartsWith(_typeToSearchString.ToLower()))
                        End If
                    Else
                        foundItem =
                            Me.Folder.Items.ToList() _
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

        Private Async Function invokeDefaultCommand(item As Item) As Task
            getMenu(Me.Folder, {item}, True)
            Await _menu.Make()
            Await _menu.InvokeCommand(_menu.DefaultId)
        End Function

        Protected Function GetIsDisplayNameTextBlockTooSmall(textBlock As TextBlock) As Boolean
            textBlock.Measure(New Size(textBlock.ActualWidth,
                If(textBlock.MaxHeight.Equals(Double.NaN), Double.PositiveInfinity, textBlock.MaxHeight)))

            Dim typeface As Typeface = New Typeface(
                textBlock.FontFamily,
                textBlock.FontStyle,
                textBlock.FontWeight,
                textBlock.FontStretch)

            Dim formattedText As FormattedText = New FormattedText(
                If(textBlock.Inlines.OfType(Of Run)().FirstOrDefault()?.Text, ""),
                System.Threading.Thread.CurrentThread.CurrentCulture,
                textBlock.FlowDirection,
                typeface,
                textBlock.FontSize,
                textBlock.Foreground,
                VisualTreeHelper.GetDpi(textBlock).PixelsPerDip)
            formattedText.MaxTextWidth = textBlock.ActualWidth + 0.25
            formattedText.TextAlignment = textBlock.TextAlignment
            formattedText.Trimming = TextTrimming.None

            Return Math.Abs(formattedText.Height - textBlock.DesiredSize.Height) > 0.5
        End Function

        Protected Overridable Sub ToggleCheckBox(checkBox As CheckBox)
            checkBox.IsChecked = Not checkBox.IsChecked
        End Sub

        Private Sub listBox_PreviewMouseMove(sender As Object, e As MouseEventArgs)
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

                _mouseOverTime = DateTime.Now

                Dim f As Func(Of Task) =
                    Async Function() As Task
                        Dim startTime As DateTime = DateTime.Now

                        Dim startOverItem As Item = overItem
                        Dim startOverTime As DateTime = _mouseOverTime
                        Dim text As String = overItem.InfoTip
                        Dim doShow As Boolean

                        UIHelper.OnUIThread(
                            Sub()
                                doShow = Me.DoShowFolderContentsInInfoTip
                                Dim textBlock As TextBlock = UIHelper.FindVisualChildren(Of TextBlock)(listBoxItem) _
                                    .FirstOrDefault(Function(b) b.Name = "PART_DisplayName" OrElse b.Tag = "PART_DisplayName")
                                If Not textBlock Is Nothing Then
                                    If Me.GetIsDisplayNameTextBlockTooSmall(textBlock) Then
                                        text = overItem.DisplayName & If(Not String.IsNullOrWhiteSpace(text), Environment.NewLine, "") & text
                                    End If
                                End If
                            End Sub)

                        If doShow Then
                            Dim textFolderSize As String = Nothing
                            If TypeOf overItem Is Folder Then
                                textFolderSize = CType(overItem, Folder).GetInfoTipFolderSizeAsync(_toolTipCancellationTokenSource.Token)
                            End If
                            If Not String.IsNullOrWhiteSpace(textFolderSize) Then
                                text &= If(Not String.IsNullOrWhiteSpace(text), Environment.NewLine, "") & textFolderSize
                            End If
                        End If

                        Await Task.Delay(Math.Max(0, 1500 - DateTime.Now.Subtract(startTime).TotalMilliseconds))

                        UIHelper.OnUIThread(
                            Sub()
                                If Not String.IsNullOrWhiteSpace(text) _
                                    AndAlso startOverItem.Equals(_mouseItemOver) _
                                    AndAlso startOverTime.Equals(_mouseOverTime) Then
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
                            End Sub)
                    End Function
                Task.Run(f)
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

        Public Sub listBox_PreviewMouseButtonDown(sender As Object, e As MouseButtonEventArgs)
            _mousePointDown = e.GetPosition(Me)
            _mouseOriginalSourceDown = e.OriginalSource
            _canOpenWithSingleClick = False

            ' this prevents a multiple selection getting replaced by the single clicked item
            If Not e.OriginalSource Is Nothing Then
                Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(e.OriginalSource)
                Dim clickedItem As Item = TryCast(listBoxItem?.DataContext, Item)
                If clickedItem Is Nothing Then
                    'Me.PART_ListBox.Focus()
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
                            ElseIf TypeOf clickedItem Is Link AndAlso TypeOf CType(clickedItem, Link).TargetItem Is Folder Then
                                CType(CType(clickedItem, Link).TargetItem, Folder).LastScrollOffset = New Point()
                                Me.Host.Folder = CType(clickedItem, Link).TargetItem
                            Else
                                Dim __ = invokeDefaultCommand(clickedItem)
                            End If
                        End Using
                    End If
                ElseIf e.LeftButton = MouseButtonState.Pressed AndAlso Not clickedItem Is Nothing Then
                    Dim checkBox As CheckBox = UIHelper.FindVisualChildren(Of CheckBox)(listBoxItem)(0)
                    Dim p As Point = Mouse.GetPosition(checkBox)
                    If Not checkBox Is Nothing AndAlso checkBox.Visibility = Visibility.Visible AndAlso p.X > 0 AndAlso p.Y > 0 _
                        AndAlso p.X < checkBox.ActualWidth AndAlso p.Y < checkBox.ActualHeight Then
                        Me.ToggleCheckBox(checkBox)
                        e.Handled = True
                    ElseIf Keyboard.Modifiers = ModifierKeys.None Then
                        If Not Me.SelectedItems Is Nothing AndAlso Me.SelectedItems.Contains(clickedItem) Then e.Handled = True
                        _canOpenWithSingleClick = True
                        _mouseItemDown = clickedItem
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
                ElseIf clickedItem Is Nothing _
                        AndAlso UIHelper.GetParentOfType(Of System.Windows.Controls.Primitives.ScrollBar)(e.OriginalSource) Is Nothing _
                        AndAlso TypeOf e.OriginalSource Is ScrollViewer Then
                    Me.SelectedItems = Nothing
                End If
            Else
                _mouseItemDown = Nothing
            End If
        End Sub

        Public Sub listBox_PreviewMouseButtonUp(sender As Object, e As MouseButtonEventArgs)
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
                        Dim __ = invokeDefaultCommand(_mouseItemDown)
                    End If
                End Using
            ElseIf Not _mouseItemDown Is Nothing AndAlso Me.IsDoubleClickToOpenItem Then
                Me.SelectedItems = {_mouseItemDown}
            End If

            _mouseItemDown = Nothing
            _canOpenWithSingleClick = False
        End Sub

        Public Sub listBox_MouseLeave(sender As Object, e As MouseEventArgs)
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

        Private Function getMenu(folder As Folder, selectedItems As IEnumerable(Of Item), isDefaultOnly As Boolean) As BaseMenu
            If folder Is Nothing Then Return Nothing

            If Not _menu Is Nothing Then
                _menu.Dispose()
            End If

            _menu = New RightClickMenu() 'New ExplorerMenu() ' 
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
                                PinnedItems.PinItem(Me.SelectedItems(0))
                            Else
                                PinnedItems.UnpinItem(Me.SelectedItems(0).Pidl)
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
            Menus.DoRename(AddressOf Me.GetItemNameCoordinates, Me.PART_Grid, listBoxItem, Me.PART_ListBox)
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

        Public Property ScrollOffset As Point
            Get
                Return GetValue(ScrollOffsetProperty)
            End Get
            Set(ByVal value As Point)
                SetCurrentValue(ScrollOffsetProperty, value)
            End Set
        End Property

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

        Public Property ExpandCollapseAllState As Boolean
            Get
                Return GetValue(ExpandCollapseAllStateProperty)
            End Get
            Set(value As Boolean)
                SetValue(ExpandCollapseAllStateProperty, value)
            End Set
        End Property

        Public Property DragViewStrategy As IDragViewStrategy
            Get
                Return _dragViewStrategy
            End Get
            Protected Set(value As IDragViewStrategy)
                _dragViewStrategy = value
            End Set
        End Property

        Protected Overridable Sub ClearBinding()
            If Not Me.PART_ListBox Is Nothing Then
                Me.PART_ListBox.ItemsSource = Nothing
            End If
        End Sub

        Protected Overridable Sub MakeBinding(folder As Folder)
            If Not Me.PART_ListBox Is Nothing Then
                Me.PART_ListBox.Visibility = Visibility.Hidden

                setGrouping(folder)

                Dim view As ListCollectionView = CollectionViewSource.GetDefaultView(folder.Items)
                Me.PART_ListBox.ItemsSource = view

                ' load items
                Dim t As Task = folder.GetItemsAsync()

                ' async because otherwise we're a tad too early
                UIHelper.OnUIThreadAsync(
                    Async Sub()
                        If Not TypeOf folder Is SearchFolder Then
                            Await t
                            'Await Task.Delay(50)

                            If Not PART_ScrollViewer Is Nothing Then
                                ' restore folder scroll position
                                Me.ScrollOffset = folder.LastScrollOffset
                                _lastScrollSize = folder.LastScrollSize
                                PART_ScrollViewer.ScrollToHorizontalOffset(If(_lastScrollSize.Width = 0, 0, Me.ScrollOffset.X * PART_ScrollViewer.ScrollableWidth / _lastScrollSize.Width))
                                PART_ScrollViewer.ScrollToVerticalOffset(If(_lastScrollSize.Height = 0, 0, Me.ScrollOffset.Y * PART_ScrollViewer.ScrollableHeight / _lastScrollSize.Height))
                            End If
                        End If

                        ' show listview
                        Me.PART_ListBox.Visibility = Visibility.Visible
                    End Sub, Threading.DispatcherPriority.ContextIdle)
            End If
        End Sub

        Protected Overridable Sub Folder_PropertyChanged(sender As Object, e As PropertyChangedEventArgs)
            Dim folder As Folder = CType(sender, Folder)

            Select Case e.PropertyName
                Case "IsRefreshingItems"
                Case "ItemsSortPropertyName", "ItemsSortDirection", "ItemsGroupByPropertyName", "ActiveView"
                    UIHelper.OnUIThread(
                        Sub()
                            If e.PropertyName = "ItemsGroupByPropertyName" Then
                                setGrouping(folder)
                            End If

                            If Not folder.IsRefreshingItems AndAlso Not Me.Folder Is Nothing AndAlso Not TypeOf folder Is SearchFolder Then
                                Dim folderViewState As FolderViewState = FolderViewState.FromViewName(folder)
                                folderViewState.SortPropertyName = folder.ItemsSortPropertyName
                                folderViewState.SortDirection = folder.ItemsSortDirection
                                folderViewState.GroupByPropertyName = folder.ItemsGroupByPropertyName
                                folderViewState.ActiveView = folder.ActiveView
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
            Me.ExpandCollapseAllState = Not isExpanded
            Me.ExpandCollapseAllState = isExpanded
            'Dim groups As IEnumerable(Of GroupItem) = UIHelper.FindVisualChildren(Of GroupItem)(PART_ListBox)
            'For Each group In groups
            '    Dim toggleButtons As IEnumerable(Of ToggleButton) = UIHelper.FindVisualChildren(Of ToggleButton)(group)
            '    If toggleButtons.Count > 0 Then
            '        toggleButtons(0).IsChecked = isExpanded
            '    End If
            'Next
        End Sub

        Private Sub folder_Items_CollectionChanged(s As Object, e As NotifyCollectionChangedEventArgs)
            updateCheckBoxSelectAll()
        End Sub

        Private Sub updateCheckBoxSelectAll()
            If Not Me.PART_CheckBoxSelectAll Is Nothing Then
                _isInternallySettingSelectAll = True
                Me.PART_CheckBoxSelectAll.IsChecked = Not Me.SelectedItems Is Nothing AndAlso Not Me.SelectedItems.Count = 0 _
                    AndAlso Me.PART_ListBox.Items.Count = Me.SelectedItems.Count
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

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
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
                oldValue.LastScrollOffset = bfv.ScrollOffset
                oldValue.LastScrollSize = bfv._lastScrollSize
                If Not bfv.PART_ScrollViewer Is Nothing Then
                    bfv.PART_ScrollViewer.ScrollToHorizontalOffset(0)
                    bfv.PART_ScrollViewer.ScrollToVerticalOffset(0)
                End If

                ' clear view binding
                bfv.ClearBinding()
            End If

            If Not e.NewValue Is Nothing Then
                Dim newValue As Folder = e.NewValue

                ' track recent/frequent folders (in a task because for some folders this might take a while)
                If Not TypeOf newValue Is SearchFolder Then
                    Dim func As Action =
                        Sub()
                            FrequentFolders.Track(newValue)
                        End Sub
                    Dim __ = Task.Run(func)
                End If

                ' set sorting and grouping
                If Not TypeOf newValue Is SearchFolder Then
                    Dim folderViewState As FolderViewState = FolderViewState.FromFolder(newValue)
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
                bfv.MakeBinding(newValue)
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