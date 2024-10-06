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
        Public Shared ReadOnly NavigationButtonsProperty As DependencyProperty = DependencyProperty.Register("NavigationButtons", GetType(List(Of StackPanel)), GetType(AddressBar), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_NavigationButtons As Border
        Private PART_ClickToEdit As Border
        Private _contextMenus As List(Of ContextMenu) = New List(Of ContextMenu)()

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(AddressBar), New FrameworkPropertyMetadata(GetType(AddressBar)))
        End Sub

        Public Sub New()
            MyBase.New()

            AddHandler Shell.Ready,
                Sub(sender As Object, e As EventArgs)
                    If Not Me.Folder Is Nothing Then
                        Me.ShowNavigationButtons(Me.Folder)
                    End If
                End Sub
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_NavigationButtons = Template.FindName("PART_NavigationButtons", Me)
            PART_ClickToEdit = Template.FindName("PART_ClickToEdit", Me)

            AddHandler Me.SizeChanged,
                Sub(s As Object, e As SizeChangedEventArgs)
                    Me.ShowNavigationButtons(Me.Folder)
                End Sub
            AddHandler PART_ClickToEdit.MouseDown,
                Sub(s As Object, e As MouseButtonEventArgs)
                    PART_NavigationButtons.Visibility = Visibility.Collapsed
                    Me.Editor.Focus()
                End Sub
        End Sub

        Protected Overrides Sub OnAfterSelect()
            If INVALID_VALUE.Equals(Me.SelectedValue) Then
                System.Media.SystemSounds.Asterisk.Play()
                Me.Cancel()
                Me.ShowNavigationButtons(Me.Folder)
            ElseIf Not Me.SelectedItem Is Nothing Then
                Shell.SetSelectedFolder(Me.SelectedItem,
                    Sub(selectedFolder As Folder)
                        If Not selectedFolder Is Nothing Then
                            Me.Folder = selectedFolder
                        Else
                            Me.Folder = Me.SelectedItem
                        End If

                        If Not PART_NavigationButtons.Visibility = Visibility.Visible Then
                            PART_NavigationButtons.Visibility = Visibility.Visible
                        End If
                    End Sub)
            Else
                Me.Cancel()
            End If
        End Sub

        Protected Overrides Sub Cancel()
            MyBase.Cancel()

            Me.ShowNavigationButtons(Me.Folder)
        End Sub

        Protected Overrides Sub OnEditorLostFocusAsync(sender As Object, e As RoutedEventArgs)
            Me.Cancel()
        End Sub

        Public Async Sub ShowNavigationButtons(folder As Folder)
            Dim buttons As List(Of StackPanel) = New List(Of StackPanel)()
            Dim totalWidth As Double = 0
            Dim buttonStyle As Style = TryFindResource("lailaShell_NavigationButtonStyle")
            Dim chevronButtonStyle As Style = TryFindResource("lailaShell_NavigationChevronButtonStyle")
            Dim moreButtonStyle As Style = TryFindResource("lailaShell_NavigationMoreButtonStyle")
            _contextMenus.Clear()

            Dim standardPanel As StackPanel = New StackPanel() With {.Orientation = Orientation.Horizontal}
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
                            Shell.SetSelectedFolder(specialFoldersMenuItem.Tag,
                                Sub(selectedFolder As Folder)
                                    If Not selectedFolder Is Nothing Then
                                        Me.Folder = selectedFolder
                                    Else
                                        Me.Folder = specialFoldersMenuItem.Tag
                                    End If
                                End Sub)
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
                        Shell.SetSelectedFolder(folderButton.Tag,
                                Sub(selectedFolder As Folder)
                                    If Not selectedFolder Is Nothing Then
                                        Me.Folder = selectedFolder
                                    Else
                                        Me.Folder = folderButton.Tag
                                    End If
                                End Sub)
                    End Sub
                panel.Children.Add(folderButton)
                Dim subFoldersButton As ToggleButton = New ToggleButton()
                If Not chevronButtonStyle Is Nothing Then subFoldersButton.Style = chevronButtonStyle
                Dim subFoldersContextMenu As ContextMenu = New ContextMenu()
                _contextMenus.Add(subFoldersContextMenu)
                subFoldersContextMenu.PlacementTarget = subFoldersButton
                subFoldersContextMenu.Placement = Primitives.PlacementMode.Bottom

                For Each subFolder In (Await currentFolder.GetItems()).Where(Function(i) TypeOf i Is Folder)
                    Dim subFoldersMenuItem As MenuItem = New MenuItem()
                    subFoldersMenuItem.Header = subFolder.DisplayName
                    subFoldersMenuItem.Tag = subFolder
                    subFoldersMenuItem.Icon = New Image() With {.Width = 16, .Height = 16, .Source = subFolder.Icon(16)}
                    AddHandler subFoldersMenuItem.Click,
                        Sub(s As Object, e As EventArgs)
                            Shell.SetSelectedFolder(subFoldersMenuItem.Tag,
                                Sub(selectedFolder As Folder)
                                    If Not selectedFolder Is Nothing Then
                                        Me.Folder = selectedFolder
                                    Else
                                        Me.Folder = subFoldersMenuItem.Tag
                                    End If
                                End Sub)
                        End Sub
                    subFoldersContextMenu.Items.Add(subFoldersMenuItem)
                Next
                If subFoldersContextMenu.Items.Count > 0 Then
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
                            Shell.SetSelectedFolder(moreMenuItem.Tag,
                                Sub(selectedFolder As Folder)
                                    If Not selectedFolder Is Nothing Then
                                        Me.Folder = selectedFolder
                                    Else
                                        Me.Folder = moreMenuItem.Tag
                                    End If
                                End Sub)
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
            If (folder Is Nothing AndAlso Me.Folder Is Nothing) OrElse folder.Equals(Me.Folder) Then ' if we're not too late with this async
                Me.NavigationButtons = buttons
                PART_NavigationButtons.Visibility = Visibility.Visible
            End If
        End Sub

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

        Public Property NavigationButtons As List(Of StackPanel)
            Get
                Return GetValue(NavigationButtonsProperty)
            End Get
            Set(ByVal value As List(Of StackPanel))
                SetCurrentValue(NavigationButtonsProperty, value)
            End Set
        End Property
    End Class
End Namespace