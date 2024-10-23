Imports Laila.AutoCompleteTextBox
Imports Laila.Shell.Helpers
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Environment
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Windows.Media

Namespace Controls
    Public Class AddressBar
        Inherits Laila.AutoCompleteTextBox.AutoCompleteTextBox

        Private Const INVALID_VALUE As String = "5e979b53-746b-4a0c-9f5f-00fdd22c91d8"

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))

        Private PART_NavigationButtonsPanel As StackPanel
        Private PART_NavigationButtons As Border
        Private PART_ClickToEdit As Border
        Private _contextMenus As List(Of ContextMenu) = New List(Of ContextMenu)()

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(AddressBar), New FrameworkPropertyMetadata(GetType(AddressBar)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_NavigationButtonsPanel = Template.FindName("PART_NavigationButtonsPanel", Me)
            PART_NavigationButtons = Template.FindName("PART_NavigationButtons", Me)
            PART_ClickToEdit = Template.FindName("PART_ClickToEdit", Me)

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    Me.ShowNavigationButtons(Me.Folder)
                End Sub
            AddHandler Me.PART_ClickToEdit.MouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    _isSettingTextInternally = True
                    Me.Text = Me.Folder.AddressBarDisplayName
                    _isSettingTextInternally = False
                    PART_NavigationButtons.Visibility = Visibility.Hidden
                    Me.PART_TextBox.IsEnabled = True
                    Me.PART_TextBox.SelectionStart = Me.Text.Length
                    Me.PART_TextBox.Focus()
                End Sub
            AddHandler Me.PART_TextBox.PreviewKeyDown,
                Sub(s As Object, e As KeyEventArgs)
                    If Not Me.IsDropDownOpen OrElse Me.IsLoadingSuggestions Then
                        Select Case e.Key
                            Case Key.Enter
                                If Not Me.IsDirty Then
                                    Me.OnItemSelected()
                                End If
                            Case Key.Escape
                                If Not Me.IsDirty Then
                                    Me.Cancel()
                                End If
                        End Select
                    End If
                End Sub
        End Sub

        Protected Overrides Sub OnItemSelected()
            Me.PART_TextBox.IsEnabled = False

            If INVALID_VALUE.Equals(Me.SelectedValue) Then
                Dim item As Item = Item.FromParsingNameDeepGetReverse(Me.PART_TextBox.Text)
                If Not item Is Nothing AndAlso TypeOf item Is Folder Then
                    Me.Folder = item
                Else
                    System.Media.SystemSounds.Asterisk.Play()
                    Me.Cancel()
                    Me.ShowNavigationButtons(Me.Folder)
                End If
            ElseIf Not Me.SelectedItem Is Nothing Then
                Dim doShow As Boolean = Me.SelectedItem.Equals(Me.Folder)
                Me.Folder = Me.SelectedItem
                If doShow Then Me.ShowNavigationButtons(Me.Folder)
            Else
                Me.Cancel()
            End If
        End Sub

        Protected Overrides Sub Cancel()
            MyBase.Cancel()

            Me.PART_TextBox.IsEnabled = False
            Me.ShowNavigationButtons(Me.Folder)
        End Sub

        Protected Overrides Sub TextBox_LostFocus(s As Object, e As RoutedEventArgs)
            Me.Cancel()
        End Sub

        Public Async Function ShowNavigationButtons(folder As Folder) As Task
            Await Task.Delay(150)

            Dim buttons As List(Of StackPanel) = New List(Of StackPanel)()
            Dim totalWidth As Double = 0
            Dim buttonStyle As Style = TryFindResource("lailaShell_NavigationButtonStyle")
            Dim chevronButtonStyle As Style = TryFindResource("lailaShell_NavigationChevronButtonStyle")
            Dim moreButtonStyle As Style = TryFindResource("lailaShell_NavigationMoreButtonStyle")
            _contextMenus.Clear()

            Dim standardPanel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal, .Focusable = False}
            Dim computerImage As Image = New Image() With {.Width = 16, .Height = 16, .VerticalAlignment = VerticalAlignment.Center}
            computerImage.Source = New ImageSourceConverter().ConvertFromInvariantString("pack://application:,,,/Laila.Shell;component/Images/monitor16.png")
            standardPanel.Children.Add(computerImage)
            Dim specialFoldersButton As ToggleButton = New ToggleButton()
            If Not chevronButtonStyle Is Nothing Then specialFoldersButton.Style = chevronButtonStyle
            Dim specialFoldersContextMenu As ContextMenu = New ContextMenu()
            _contextMenus.Add(specialFoldersContextMenu)
            specialFoldersContextMenu.PlacementTarget = specialFoldersButton
            specialFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
            For Each specialFolder In Shell.SpecialFolders.Values.OrderBy(Function(f) f.DisplayName)
                Dim specialFoldersMenuItem As MenuItem = New MenuItem()
                specialFoldersMenuItem.Header = specialFolder.DisplayName
                specialFoldersMenuItem.Tag = specialFolder
                specialFoldersMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = specialFolder.Icon(16)}
                AddHandler specialFoldersMenuItem.Click,
                        Sub(s As Object, e As EventArgs)
                            Me.Folder = specialFoldersMenuItem.Tag
                        End Sub
                specialFoldersContextMenu.Items.Add(specialFoldersMenuItem)
            Next
            AddHandler specialFoldersButton.Checked,
                            Sub(s As Object, e As EventArgs)
                                specialFoldersContextMenu.IsOpen = True
                            End Sub
            AddHandler specialFoldersContextMenu.Closed,
                            Sub(s As Object, e As EventArgs)
                                specialFoldersButton.IsChecked = False
                            End Sub
            standardPanel.Children.Add(specialFoldersButton)
            Dim moreButton As ToggleButton = New ToggleButton()
            If Not moreButtonStyle Is Nothing Then moreButton.Style = moreButtonStyle
            standardPanel.Children.Add(moreButton)
            standardPanel.Measure(New Size(1000, 1000))
            totalWidth += standardPanel.DesiredSize.Width

            Dim currentFolder As Folder = folder
            While Not currentFolder Is Nothing
                Dim panel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal}
                Dim folderButton As Button = New Button()
                If Not buttonStyle Is Nothing Then folderButton.Style = buttonStyle
                folderButton.Content = currentFolder.DisplayName
                folderButton.Tag = currentFolder
                AddHandler folderButton.Click,
                    Sub(s As Object, e As EventArgs)
                        Me.Folder = folderButton.Tag
                    End Sub
                panel.Children.Add(folderButton)
                Dim subFoldersButton As ToggleButton = New ToggleButton()
                If Not chevronButtonStyle Is Nothing Then subFoldersButton.Style = chevronButtonStyle
                Dim subFoldersContextMenu As ContextMenu = New ContextMenu()
                subFoldersContextMenu.PlacementTarget = subFoldersButton
                subFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom
                _contextMenus.Add(subFoldersContextMenu)

                For Each subFolder In (Await currentFolder.GetItemsAsync()).Where(Function(i) TypeOf i Is Folder)
                    Dim subFoldersMenuItem As MenuItem = New MenuItem()
                    subFoldersMenuItem.Header = subFolder.DisplayName
                    subFoldersMenuItem.Tag = subFolder
                    subFoldersMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = subFolder.Icon(16)}
                    AddHandler subFoldersMenuItem.Click,
                                Sub(s As Object, e As EventArgs)
                                    Me.Folder = subFoldersMenuItem.Tag
                                End Sub
                    subFoldersContextMenu.Items.Add(subFoldersMenuItem)
                Next
                If subFoldersContextMenu.Items.Count > 0 Then
                    AddHandler subFoldersContextMenu.Opened,
                                Sub(s As Object, e As EventArgs)
                                    subFoldersButton.IsChecked = True
                                End Sub
                    AddHandler subFoldersButton.Checked,
                                Sub(s As Object, e As EventArgs)
                                    subFoldersContextMenu.IsOpen = True
                                End Sub
                    panel.Children.Add(subFoldersButton)
                    AddHandler subFoldersContextMenu.Closed,
                                Sub(s As Object, e As EventArgs)
                                    subFoldersButton.IsChecked = False
                                End Sub
                End If

                panel.Measure(New Size(1000, 1000))
                If totalWidth + panel.DesiredSize.Width < Me.ActualWidth - 30 Then
                    totalWidth += panel.DesiredSize.Width
                    buttons.Add(panel)
                Else
                    Exit While
                End If
                currentFolder = currentFolder.LogicalParent
            End While

            If Not currentFolder Is Nothing Then
                Dim moreContextMenu As ContextMenu = New ContextMenu()
                moreContextMenu.PlacementTarget = moreButton
                moreContextMenu.Placement = Primitives.PlacementMode.Bottom
                _contextMenus.Add(moreContextMenu)
                While Not currentFolder Is Nothing
                    Dim moreMenuItem As MenuItem = New MenuItem()
                    moreMenuItem.Header = currentFolder.DisplayName
                    moreMenuItem.Tag = currentFolder
                    moreMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = currentFolder.Icon(16)}
                    AddHandler moreMenuItem.Click,
                        Sub(s As Object, e As EventArgs)
                            Me.Folder = moreMenuItem.Tag
                        End Sub
                    moreContextMenu.Items.Add(moreMenuItem)
                    currentFolder = currentFolder.LogicalParent
                End While
                AddHandler moreButton.Checked,
                    Sub(s As Object, e As EventArgs)
                        moreContextMenu.IsOpen = True
                    End Sub
                AddHandler moreContextMenu.Closed,
                         Sub(s As Object, e As EventArgs)
                             moreButton.IsChecked = False
                         End Sub
            Else
                moreButton.Visibility = Visibility.Collapsed
            End If

            buttons.Add(standardPanel)
            buttons.Reverse()
            If (folder Is Nothing AndAlso Me.Folder Is Nothing) OrElse (Not folder Is Nothing AndAlso folder.Equals(Me.Folder)) Then ' if we're not too late with this async
                Me.PART_NavigationButtonsPanel.Children.Clear()
                For Each button In buttons
                    Me.PART_NavigationButtonsPanel.Children.Add(button)
                Next
                Me.PART_NavigationButtons.Visibility = Visibility.Visible
            End If
        End Function

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(ByVal value As Folder)
                SetCurrentValue(FolderProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim ab As AddressBar = TryCast(d, AddressBar)
            ab.SelectedItem = e.NewValue
            ab.ShowNavigationButtons(e.NewValue)
        End Sub
    End Class
End Namespace