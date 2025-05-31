Imports System.Reflection
Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Controls.Primitives
Imports System.Windows.Data
Imports System.Windows.Input
Imports System.Windows.Media
Imports System.Windows.Threading
Imports Laila.Shell.Controls.Parts
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class AutoCompleteTextBox
        Inherits BaseControl

        Public Shared ReadOnly AllowFreeTextProperty As DependencyProperty = DependencyProperty.Register("AllowFreeText", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ShowDropDownButtonProperty As DependencyProperty = DependencyProperty.Register("ShowDropDownButton", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ProviderProperty As DependencyProperty = DependencyProperty.Register("Provider", GetType(ISuggestionProviderSyncOrAsync), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnProviderChanged))
        Public Shared ReadOnly WatermarkProperty As DependencyProperty = DependencyProperty.Register("Watermark", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MinCharsProperty As DependencyProperty = DependencyProperty.Register("MinChars", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(3, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly MaxLengthProperty As DependencyProperty = DependencyProperty.Register("MaxLength", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(0, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DelayProperty As DependencyProperty = DependencyProperty.Register("Delay", GetType(Integer), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(500, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DisplayMemberProperty As DependencyProperty = DependencyProperty.Register("DisplayMember", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IconProperty As DependencyProperty = DependencyProperty.Register("Icon", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly InvalidValueProperty As DependencyProperty = DependencyProperty.Register("InvalidValue", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Integer.MinValue, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDropDownOpenProperty As DependencyProperty = DependencyProperty.Register("IsDropDownOpen", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsBalloonOpenProperty As DependencyProperty = DependencyProperty.Register("IsBalloonOpen", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsReadOnlyProperty As DependencyProperty = DependencyProperty.Register("IsReadOnly", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsToolTipEnabledProperty As DependencyProperty = DependencyProperty.Register("IsToolTipEnabled", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ItemTemplateProperty As DependencyProperty = DependencyProperty.Register("ItemTemplate", GetType(DataTemplate), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly LoadingContentProperty As DependencyProperty = DependencyProperty.Register("LoadingContent", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedValuePathProperty As DependencyProperty = DependencyProperty.Register("SelectedValuePath", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly TextProperty As DependencyProperty = DependencyProperty.Register("Text", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(String.Empty, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly SelectedItemProperty As DependencyProperty = DependencyProperty.Register("SelectedItem", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemChanged))
        Public Shared ReadOnly SelectedValueProperty As DependencyProperty = DependencyProperty.Register("SelectedValue", GetType(Object), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedValueChanged))
        Public Shared ReadOnly IsInvalidProperty As DependencyProperty = DependencyProperty.Register("IsInvalid", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsLoadingSuggestionsProperty As DependencyProperty = DependencyProperty.Register("IsLoadingSuggestions", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly ErrorMessageProperty As DependencyProperty = DependencyProperty.Register("ErrorMessage", GetType(String), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDirtyProperty As DependencyProperty = DependencyProperty.Register("IsDirty", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared Shadows ReadOnly IsTabStopProperty As DependencyProperty = DependencyProperty.Register("IsTabStop", GetType(Boolean), GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(True, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Protected PART_TextBox As TextBox
        Protected PART_ListBox As ListBox
        Protected PART_Popup As Popup
        Protected PART_DropDownButton As ToggleButton
        Protected _isSettingValueAfterUserInput As Boolean
        Protected _isSettingTextInternally As Boolean
        Protected _isGettingItem As Boolean
        Private _timer As DispatcherTimer
        Private _cancelValue As Object
        Private _cancelText As String
        Private _cancelSelectText As String
        Private _cancellationTokenSource As CancellationTokenSource

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(AutoCompleteTextBox), New FrameworkPropertyMetadata(GetType(AutoCompleteTextBox)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            Debug.WriteLine("Public Sub OnApplyTemplate()")
            MyBase.OnApplyTemplate()

            MyBase.IsTabStop = False

            Me.PART_TextBox = Me.Template.FindName("PART_TextBox", Me)
            Me.PART_ListBox = Me.Template.FindName("PART_ListBox", Me)
            Me.PART_Popup = Me.Template.FindName("PART_Popup", Me)
            Me.PART_DropDownButton = Me.Template.FindName("PART_DropDownButton", Me)

            AddHandler Me.PART_TextBox.TextChanged, AddressOf TextBox_TextChanged
            AddHandler Me.PART_TextBox.PreviewKeyDown, AddressOf TextBox_KeyDown
            AddHandler Me.PART_TextBox.LostFocus, AddressOf TextBox_LostFocus

            AddHandler Me.PART_ListBox.PreviewMouseDown, AddressOf listBox_PreviewMouseDown

            AddHandler Me.PART_Popup.Closed,
            Sub(s As Object, e As EventArgs)
                Debug.WriteLine("Me.PART_Popup.Closed")
                Me.PART_TextBox.SetBinding(TextBox.TextProperty, New Binding("Text") With {.Source = Me, .UpdateSourceTrigger = UpdateSourceTrigger.PropertyChanged})
                Me.IsDropDownOpen = False
            End Sub

            AddHandler Me.PART_DropDownButton.Checked,
            Sub(s As Object, e As RoutedEventArgs)
                Debug.WriteLine("Me.PART_DropDownButton.Checked")
                If Not Me.IsLoadingSuggestions Then
                    Me.PART_TextBox.Focus()
                    Dim __ = displaySuggestions("", True)
                End If
            End Sub

            AddHandler Me.IsEnabledChanged,
            Sub(s As Object, e As DependencyPropertyChangedEventArgs)
                Debug.WriteLine("Me.IsEnabledChanged")
                If Not Me.IsEnabled AndAlso Me.IsDropDownOpen Then
                    Me.IsDropDownOpen = False
                End If
            End Sub

            AddHandler Me.GotFocus,
            Sub(s As Object, e As EventArgs)
                Me.FocusChild()
            End Sub
        End Sub

        Public Overridable Sub FocusChild()
            If Not Me.PART_TextBox Is Nothing Then
                Me.PART_TextBox.Focus()
            End If
        End Sub

        Protected Overridable Sub TextBox_TextChanged(s As Object, e As TextChangedEventArgs)
            Debug.WriteLine("Protected Sub TextBox_TextChanged")
            If Me.PART_TextBox.Text Is Nothing OrElse Me.PART_TextBox.Text.Length = 0 Then
                Me.IsDropDownOpen = False
                killTimer()
                Me.SelectedItem = Nothing
                Me.SelectedValue = Nothing
                Me.IsDirty = False
            ElseIf Not _isSettingTextInternally Then
                If Me.PART_TextBox.Text.Length >= Me.MinChars Then
                    resetTimer()
                Else
                    Me.IsDropDownOpen = False
                End If
                Me.IsDirty = True
            End If
        End Sub

        Protected Overridable Sub TextBox_KeyDown(s As Object, e As KeyEventArgs)
            Debug.WriteLine("Protected Sub TextBox_KeyDown")
            If Me.IsDropDownOpen AndAlso Not Me.IsLoadingSuggestions Then
                Select Case e.Key
                    Case Key.Down
                        If Me.PART_ListBox.SelectedIndex = -1 Then
                            Me.PART_ListBox.SelectedIndex = 0
                        ElseIf Me.PART_ListBox.SelectedIndex = Me.PART_ListBox.Items.Count - 1 Then
                            Me.PART_ListBox.SelectedIndex = 0
                        Else
                            Me.PART_ListBox.SelectedIndex += 1
                        End If
                        selectionChanged()
                        e.Handled = True
                    Case Key.Up
                        If Me.PART_ListBox.SelectedIndex = -1 Then
                            Me.PART_ListBox.SelectedIndex = Me.PART_ListBox.Items.Count - 1
                        ElseIf Me.PART_ListBox.SelectedIndex = 0 Then
                            Me.PART_ListBox.SelectedIndex = Me.PART_ListBox.Items.Count - 1
                        Else
                            Me.PART_ListBox.SelectedIndex -= 1
                        End If
                        selectionChanged()
                        e.Handled = True
                    Case Key.Enter
                        commitSelectionFromList()
                    Case Key.Escape
                        e.Handled = True
                        cancelSelectFromList()
                    Case Key.Tab
                        commitSelectionFromList()
                    Case Key.PageDown
                        Dim items As Integer = Me.PART_ListBox.ActualHeight / CType(Me.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem).ActualHeight - 1
                        If Me.PART_ListBox.SelectedIndex = -1 Then Me.PART_ListBox.SelectedIndex = 0
                        For i = 1 To items
                            If Me.PART_ListBox.SelectedIndex < Me.PART_ListBox.Items.Count - 1 Then
                                Me.PART_ListBox.SelectedIndex += 1
                            End If
                        Next
                        selectionChanged()
                        e.Handled = True
                    Case Key.PageUp
                        Dim items As Integer = Me.PART_ListBox.ActualHeight / CType(Me.PART_ListBox.ItemContainerGenerator.ContainerFromIndex(0), ListBoxItem).ActualHeight - 1
                        If Me.PART_ListBox.SelectedIndex = -1 Then Me.PART_ListBox.SelectedIndex = 0
                        For i = 1 To items
                            If Me.PART_ListBox.SelectedIndex > 0 Then
                                Me.PART_ListBox.SelectedIndex -= 1
                            End If
                        Next
                        selectionChanged()
                        e.Handled = True
                End Select
            Else
                Select Case e.Key
                    Case Key.System
                        If Keyboard.Modifiers.HasFlag(ModifierKeys.Alt) AndAlso e.SystemKey = Key.Down Then
                            Dim __ = displaySuggestions("", True)
                        End If
                    Case Key.Enter
                        If Me.IsDirty Then
                            _isSettingValueAfterUserInput = True
                            If TypeOf Me.Provider Is ISuggestionProvider Then
                                tryGetItemSync(Me.Text)
                            ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                                Me.WaitSyncForAsync(
                                Async Function() As Task
                                    Await tryGetItemAsync(Me.SelectedValue, Me.Text)
                                End Function)
                            End If
                            _isSettingValueAfterUserInput = False
                        End If
                    Case Key.Escape
                        If Me.IsDirty Then
                            Me.Cancel()
                        End If
                End Select
            End If
        End Sub

        Protected Overridable Sub TextBox_LostFocus(s As Object, e As RoutedEventArgs)
            Debug.WriteLine("Protected Sub TextBox_LostFocus")
            If Me.IsDirty Then
                _isSettingValueAfterUserInput = True
                If TypeOf Me.Provider Is ISuggestionProvider Then
                    tryGetItemSync(Me.Text)
                ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                    Me.WaitSyncForAsync(
                    Async Function() As Task
                        Await tryGetItemAsync(Me.SelectedValue, Me.Text)
                    End Function)
                End If
                _isSettingValueAfterUserInput = False
            End If
        End Sub

        Private Async Function displaySuggestions(text As String, doSelectCurrent As Boolean) As Task
            Debug.WriteLine("Private Function displaySuggestions 1")
            If Not _cancellationTokenSource Is Nothing Then
                _cancellationTokenSource.Cancel()
            End If
            _cancellationTokenSource = New CancellationTokenSource()
            Await displaySuggestions(text, doSelectCurrent, _cancellationTokenSource.Token)
        End Function

        Private Async Function displaySuggestions(text As String, doSelectCurrent As Boolean, cancellationToken As CancellationToken) As Task
            Debug.WriteLine("Private Function displaySuggestions 2")
            killTimer()
            Me.PART_ListBox.ItemsSource = Nothing
            Me.IsLoadingSuggestions = True
            Me.IsDropDownOpen = True
            Me.IsBalloonOpen = False

            If Me.IsEnabled Then
                Dim list As IEnumerable

                Try
                    If cancellationToken.IsCancellationRequested Then Return

                    If TypeOf Me.Provider Is ISuggestionProvider Then
                        Dim provider As ISuggestionProvider = Me.Provider
                        Dim func As Func(Of Task(Of IEnumerable)) =
                            Function() As Task(Of IEnumerable)
                                Return Task.FromResult(provider.GetSuggestions(text))
                            End Function
                        list = Await Task.Run(func)
                    ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                        list = Await CType(Me.Provider, ISuggestionProviderAsync).GetSuggestions(text)
                    Else
                        Throw New InvalidOperationException(String.Format("Unknown suggestionprovider type '{0}'.", Me.Provider.GetType()))
                    End If

                    If cancellationToken.IsCancellationRequested Then Return

                    If Me.IsKeyboardFocusWithin AndAlso Not list Is Nothing AndAlso list.GetEnumerator().MoveNext() Then
                        If New List(Of Object)(list).Count > 350 Then
                            ' wait for dropdown to open
                            Await Task.Delay(250)
                        End If

                        Me.PART_ListBox.ItemsSource = list
                        _cancelSelectText = Me.Text
                        BindingOperations.ClearBinding(Me.PART_TextBox, TextBox.TextProperty)

                        If doSelectCurrent AndAlso Not Me.SelectedValue Is Nothing AndAlso Not getIsSelectedValueInvalid() Then
                            Me.PART_ListBox.SelectedItem = list.Cast(Of Object).FirstOrDefault(Function(i) _
                            Me.SelectedValue.Equals(getValueFromItem(i)))
                            If Not Me.PART_ListBox.SelectedItem Is Nothing Then
                                Me.PART_ListBox.ScrollIntoView(Me.PART_ListBox.SelectedItem)
                            End If
                        End If
                    Else
                        Me.IsDropDownOpen = False
                    End If
                Catch ex As Exception
                    OnLoadingException(ex)
                Finally
                    Me.IsLoadingSuggestions = False
                End Try
            Else
                Me.IsLoadingSuggestions = False
            End If
        End Function

        Private Sub tryGetItemSync(Optional callback As Action = Nothing)
            Debug.WriteLine("Private Sub tryGetItemSync 1")
            _isGettingItem = True

            If Not getIsSelectedValueMatchesSelectedItem() Then
                If Not Me.Provider Is Nothing AndAlso Not Me.SelectedValue Is Nothing AndAlso Not getIsSelectedValueInvalid() Then
                    tryGetItemSync(String.Format("GetById://{0}", Me.SelectedValue), callback)
                    Return
                ElseIf Not Me.SelectedItem Is Nothing Then
                    Me.SelectedItem = Nothing
                End If
            End If

            If Not callback Is Nothing Then
                callback()
            End If

            _isGettingItem = False
        End Sub

        Private Sub tryGetItemSync(filter As String, Optional callback As Action = Nothing)
            Debug.WriteLine("Private Sub tryGetItemSync 2")
            If Not String.IsNullOrWhiteSpace(filter) Then
                Dim list As IEnumerable

                Try
                    _isGettingItem = True

                    If Not _cancellationTokenSource Is Nothing Then
                        _cancellationTokenSource.Cancel()
                        _cancellationTokenSource = Nothing
                    End If
                    killTimer()

                    If TypeOf Me.Provider Is ISuggestionProvider Then
                        list = CType(Provider, ISuggestionProvider).GetSuggestions(filter)
                    Else
                        Throw New InvalidOperationException(String.Format("Invalid suggestionprovider type '{0}'.", Me.Provider.GetType()))
                    End If

                    Me.IsDirty = False
                    Me.IsDropDownOpen = False

                    If Not list Is Nothing AndAlso list.GetEnumerator().MoveNext() Then
                        If Not list(0).Equals(Me.SelectedItem) Then
                            Me.SelectedItem = list(0)
                        Else
                            Me.OnSelectedItemChanged()
                        End If
                    Else
                        Application.Current.Dispatcher.InvokeAsync(
                        Sub()
                            Me.SelectedValue = Me.InvalidValue
                        End Sub)
                    End If

                    If Not callback Is Nothing Then
                        callback()
                    End If

                    If _isSettingValueAfterUserInput Then Me.OnItemSelected()
                Catch ex As Exception
                    OnLoadingException(ex)
                Finally
                    _isGettingItem = False
                End Try
            End If
        End Sub

        Private Async Function tryGetItemAsync(initialSelectedValue As Object, Optional callback As Action = Nothing) As Task
            Debug.WriteLine("Private Function tryGetItemAsync 1")
            _isGettingItem = True

            If Not getIsSelectedValueMatchesSelectedItem() Then
                If Not Me.Provider Is Nothing AndAlso Not Me.SelectedValue Is Nothing AndAlso Not getIsSelectedValueInvalid() Then
                    Await tryGetItemAsync(initialSelectedValue, String.Format("GetById://{0}", Me.SelectedValue), callback)
                    Return
                ElseIf Not Me.SelectedItem Is Nothing Then
                    Me.SelectedItem = Nothing
                End If
            End If

            If Not callback Is Nothing Then
                callback()
            End If

            _isGettingItem = False
        End Function

        Private Async Function tryGetItemAsync(initialSelectedValue As Object, filter As String, Optional callback As Action = Nothing) As Task
            Debug.WriteLine("Private Function tryGetItemAsync")
            If Not String.IsNullOrWhiteSpace(filter) Then
                Dim list As IEnumerable

                Try
                    _isGettingItem = True

                    If Not _cancellationTokenSource Is Nothing Then
                        _cancellationTokenSource.Cancel()
                        _cancellationTokenSource = Nothing
                    End If
                    killTimer()

                    If TypeOf Me.Provider Is ISuggestionProviderAsync Then
                        list = Await CType(Me.Provider, ISuggestionProviderAsync).GetSuggestions(filter)
                    Else
                        Throw New InvalidOperationException(String.Format("Invalid suggestionprovider type '{0}'.", Me.Provider.GetType()))
                    End If

                    Me.IsDirty = False
                    Me.IsDropDownOpen = False

                    Application.Current.Dispatcher.Invoke(
                    Sub()
                        If EqualityComparer(Of Object).Default.Equals(Me.SelectedValue, initialSelectedValue) Then
                            If Not list Is Nothing AndAlso list.GetEnumerator().MoveNext() Then
                                If Not list(0).Equals(Me.SelectedItem) Then
                                    Me.SelectedItem = list(0)
                                Else
                                    Me.OnSelectedItemChanged()
                                End If
                            Else
                                Me.SelectedValue = Me.InvalidValue
                            End If

                            If Not callback Is Nothing Then
                                callback()
                            End If

                            If _isSettingValueAfterUserInput Then Me.OnItemSelected()
                        End If
                    End Sub)
                Catch ex As Exception
                    OnLoadingException(ex)
                Finally
                    _isGettingItem = False
                End Try
            End If
        End Function

        Protected Overridable Sub OnLoadingException(ex As Exception)
            Debug.WriteLine("Protected Sub OnLoadingException")
            Me.IsDropDownOpen = False
            Me.ErrorMessage = ex.Message & If(Not ex.InnerException Is Nothing, vbCrLf & ex.InnerException.Message, "")
            Me.IsBalloonOpen = True
        End Sub

        Private Sub killTimer()
            Debug.WriteLine("Private Sub killTimer()")
            If Not _timer Is Nothing Then
                _timer.Stop()
            End If
        End Sub

        Private Sub resetTimer()
            Debug.WriteLine("Private Sub resetTimer()")
            killTimer()
            If _timer Is Nothing Then
                _timer = New DispatcherTimer()
                _timer.Interval = TimeSpan.FromMilliseconds(Me.Delay)
                AddHandler _timer.Tick,
                Async Sub(s As Object, e As EventArgs)
                    Await displaySuggestions(Me.Text, False)
                End Sub
            End If
            _timer.Start()
        End Sub

        Protected Overridable Sub Cancel()
            Debug.WriteLine("Protected Sub Cancel()")
            If Me.IsDropDownOpen Then
                Me.IsDropDownOpen = False
                _isSettingTextInternally = True
                Me.Text = _cancelSelectText
                _isSettingTextInternally = False
            Else
                _isSettingTextInternally = True
                Me.Text = _cancelText
                _isSettingTextInternally = False
                Me.SelectedValue = _cancelValue
                Me.IsDirty = False
            End If
            Me.PART_TextBox.SelectionStart = If(Not Me.Text Is Nothing, Me.Text.Length, 0)
            Me.OnCancelled()
        End Sub

        Protected Overridable Sub OnCancelled()
            Debug.WriteLine("Protected Sub OnCancelled()")
        End Sub

        Protected Overridable Sub OnItemSelected()
            Debug.WriteLine("Protected Sub OnItemSelected()")
        End Sub

        Private Sub cancelSelectFromList()
            Debug.WriteLine("Private Sub cancelSelectFromList()")
            Me.Cancel()
        End Sub

        Private Sub commitSelectionFromList()
            Debug.WriteLine("Private Sub commitSelectionFromList()")
            Me.IsDropDownOpen = False
            Me.IsDirty = False
            _isSettingValueAfterUserInput = True
            If Not Me.PART_ListBox.SelectedItem Is Nothing Then
                Me.SelectedItem = Me.PART_ListBox.SelectedItem
            Else
                If TypeOf Me.Provider Is ISuggestionProvider Then
                    tryGetItemSync(Me.Text)
                ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                    Me.WaitSyncForAsync(
                    Async Function() As Task
                        Await tryGetItemAsync(Me.SelectedValue, Me.Text)
                    End Function)
                End If
            End If
            _isSettingValueAfterUserInput = False
            Me.PART_ListBox.SelectedItem = Nothing
            Me.PART_TextBox.Focus()
            Me.PART_TextBox.SelectAll()
            Me.OnItemSelected()
        End Sub

        Private Sub selectionChanged()
            Debug.WriteLine("Private Sub selectionChanged()")
            _isSettingTextInternally = True
            Me.PART_TextBox.Text = getTextFromItem(Me.PART_ListBox.SelectedItem)
            _isSettingTextInternally = False
            Me.PART_TextBox.SelectionStart = PART_TextBox.Text.Length
            Me.PART_TextBox.SelectionLength = 0
            Me.PART_ListBox.ScrollIntoView(Me.PART_ListBox.SelectedItem)
        End Sub

        Private Sub listBox_PreviewMouseDown(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            Debug.WriteLine("Private Sub listBox_PreviewMouseDown")
            Dim result As HitTestResult = VisualTreeHelper.HitTest(Me.PART_ListBox, e.GetPosition(Me.PART_ListBox))
            If Not result Is Nothing Then
                Dim listBoxItem As ListBoxItem = UIHelper.GetParentOfType(Of ListBoxItem)(result.VisualHit, Me.PART_ListBox)
                If Not listBoxItem Is Nothing Then
                    listBoxItem.IsSelected = True
                    ' capture the mouse because we're previewing mousedown and we don't want the click to go to the control below
                    Mouse.Capture(Me.PART_ListBox, CaptureMode.SubTree)
                    selectionChanged()
                    commitSelectionFromList()
                End If
            End If
        End Sub

        Private Function getIsSelectedValueMatchesSelectedItem() As Boolean
            Debug.WriteLine("Private Function getIsSelectedValueMatchesSelectedItem()")
            If Me.SelectedItem Is Nothing AndAlso getIsSelectedValueInvalid() Then
                Return True
            Else
                ' this says string "3" equals int (3) which is what we need for backwards compatibility
                ' and for not getting in an endless loop when the two properties mismatch types
                Return (getValueFromItem(Me.SelectedItem) Is Nothing AndAlso Me.SelectedValue Is Nothing) _
                OrElse (Not getValueFromItem(Me.SelectedItem) Is Nothing AndAlso Not Me.SelectedValue Is Nothing _
                        AndAlso getValueFromItem(Me.SelectedItem) = Me.SelectedValue)
            End If
        End Function

        Private Function getIsSelectedValueInvalid() As Boolean
            Debug.WriteLine("Private Function getIsSelectedValueInvalid()")
            Return EqualityComparer(Of Object).Default.Equals(Me.SelectedValue, Me.InvalidValue)
        End Function

        Private Function getTextFromItem(item As Object) As String
            Debug.WriteLine("Private Function getTextFromItem")
            If item Is Nothing Then
                Return String.Empty
            ElseIf String.IsNullOrEmpty(Me.DisplayMember) Then
                Return item.ToString()
            Else
                Dim p As PropertyInfo = item.GetType().GetProperty(Me.DisplayMember)
                If Not p Is Nothing Then
                    Return p.GetValue(item, Nothing)
                Else
                    Return Nothing
                End If
            End If
        End Function

        Private Function getValueFromItem(item As Object) As Object
            Debug.WriteLine("Private Function getValueFromItem")
            If Not item Is Nothing AndAlso Not String.IsNullOrWhiteSpace(Me.SelectedValuePath) Then
                Dim p As PropertyInfo = item.GetType().GetProperty(Me.SelectedValuePath)
                If Not p Is Nothing Then
                    Return p.GetValue(item, Nothing)
                Else
                    Return Nothing
                End If
            Else
                Return Nothing
            End If
        End Function

        Private Sub WaitSyncForAsync(func As Func(Of Task))
            Debug.WriteLine("Private Sub WaitSyncForAsync")
            Dim result As Boolean
            Application.Current.Dispatcher.Invoke(
            Async Function() As Task
                Await func()
                result = True
            End Function)
            While Not result
                Application.Current.Dispatcher.Invoke(
                Sub()
                End Sub, Threading.DispatcherPriority.ContextIdle)
            End While
        End Sub

        Protected Friend Overridable Sub OnSelectedItemChanged()
            Debug.WriteLine("Protected Friend Sub OnSelectedItemChanged()")
            ' hide tooltip?
            If Me.SelectedItem Is Nothing AndAlso Not Me.ToolTip Is Nothing AndAlso TypeOf (Me.ToolTip) Is ToolTip Then
                CType(Me.ToolTip, ToolTip).IsOpen = False
            End If

            ' set display text
            If Not Me.SelectedItem Is Nothing Then
                _isSettingTextInternally = True
                Me.Text = getTextFromItem(Me.SelectedItem)
                _isSettingTextInternally = False
                If Not Me.PART_TextBox Is Nothing Then
                    Me.PART_TextBox.SelectAll()
                End If
            ElseIf Not Me.AllowFreeText AndAlso Not _isSettingValueAfterUserInput Then
                Me.Text = Nothing
            End If

            ' set value to match item
            If Not getIsSelectedValueMatchesSelectedItem() Then
                Me.SelectedValue = getValueFromItem(Me.SelectedItem)
            End If
        End Sub

        Shared Sub OnSelectedItemChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Debug.WriteLine("Shared Sub OnSelectedItemChanged")
            Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
            act.OnSelectedItemChanged()
        End Sub

        Protected Friend Overridable Sub OnSelectedValueChanged(oldValue As Object, newValue As Object)
            Debug.WriteLine("Protected Friend Sub OnSelectedValueChanged")
            Dim callback As Action =
            Sub()
                If (EqualityComparer(Of Object).Default.Equals(oldValue, Me.InvalidValue) AndAlso newValue Is Nothing) _
                    OrElse (EqualityComparer(Of Object).Default.Equals(newValue, Me.InvalidValue) AndAlso oldValue Is Nothing) Then
                    Me.OnSelectedItemChanged()
                End If

                Me.IsInvalid = Not Me.AllowFreeText AndAlso getIsSelectedValueInvalid()

                _cancelText = Me.Text
                _cancelValue = Me.SelectedValue
            End Sub

            If TypeOf Me.Provider Is ISuggestionProvider Then
                tryGetItemSync(callback)
            ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                Dim __ = tryGetItemAsync(newValue, callback)
            End If
        End Sub

        Shared Sub OnSelectedValueChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Debug.WriteLine("Shared Sub OnSelectedValueChanged")
            Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
            act.OnSelectedValueChanged(e.OldValue, e.NewValue)
        End Sub

        Friend Sub OnProviderChanged()
            Debug.WriteLine("Friend Sub OnProviderChanged()")
            If TypeOf Me.Provider Is ISuggestionProvider Then
                tryGetItemSync()
            ElseIf TypeOf Me.Provider Is ISuggestionProviderAsync Then
                Dim __ = tryGetItemAsync(Me.SelectedValue)
            End If
        End Sub

        Shared Sub OnProviderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Debug.WriteLine("Shared Sub OnProviderChanged")
            Dim act As AutoCompleteTextBox = TryCast(d, AutoCompleteTextBox)
            act.OnProviderChanged()
        End Sub

        Public Property AllowFreeText As Boolean
            Get
                Return GetValue(AllowFreeTextProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(AllowFreeTextProperty, value)
            End Set
        End Property

        Public Property Provider As ISuggestionProviderSyncOrAsync
            Get
                Return GetValue(ProviderProperty)
            End Get
            Set(value As ISuggestionProviderSyncOrAsync)
                SetCurrentValue(ProviderProperty, value)
            End Set
        End Property

        Public Property MinChars As Integer
            Get
                Return GetValue(MinCharsProperty)
            End Get
            Set(value As Integer)
                SetCurrentValue(MinCharsProperty, value)
            End Set
        End Property

        Public Property Delay As Integer
            Get
                Return GetValue(DelayProperty)
            End Get
            Set(value As Integer)
                SetCurrentValue(DelayProperty, value)
            End Set
        End Property

        Public Property DisplayMember As String
            Get
                Return GetValue(DisplayMemberProperty)
            End Get
            Set(value As String)
                SetCurrentValue(DisplayMemberProperty, value)
            End Set
        End Property

        Public Property InvalidValue As Object
            Get
                Return GetValue(InvalidValueProperty)
            End Get
            Set(value As Object)
                SetCurrentValue(InvalidValueProperty, value)
            End Set
        End Property

        Public Property IsDropDownOpen As Boolean
            Get
                Return GetValue(IsDropDownOpenProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsDropDownOpenProperty, value)
            End Set
        End Property

        Public Property IsBalloonOpen As Boolean
            Get
                Return GetValue(IsBalloonOpenProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsBalloonOpenProperty, value)
            End Set
        End Property

        Public Property SelectedValuePath As String
            Get
                Return GetValue(SelectedValuePathProperty)
            End Get
            Set(value As String)
                SetCurrentValue(SelectedValuePathProperty, value)
            End Set
        End Property

        Public Property Text As String
            Get
                Return GetValue(TextProperty)
            End Get
            Set(value As String)
                SetCurrentValue(TextProperty, value)
            End Set
        End Property

        Public Overloads Property SelectedItem As Object
            Get
                Return GetValue(SelectedItemProperty)
            End Get
            Set(value As Object)
                SetCurrentValue(SelectedItemProperty, value)
            End Set
        End Property

        Public Overloads Property SelectedValue As Object
            Get
                Return GetValue(SelectedValueProperty)
            End Get
            Set(value As Object)
                SetCurrentValue(SelectedValueProperty, value)
            End Set
        End Property

        Public Property IsInvalid As Boolean
            Get
                Return GetValue(IsInvalidProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsInvalidProperty, value)
            End Set
        End Property

        Public Property IsLoadingSuggestions As Boolean
            Get
                Return GetValue(IsLoadingSuggestionsProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsLoadingSuggestionsProperty, value)
            End Set
        End Property

        Public Property ErrorMessage As String
            Get
                Return GetValue(ErrorMessageProperty)
            End Get
            Set(value As String)
                SetCurrentValue(ErrorMessageProperty, value)
            End Set
        End Property

        Public Property IsDirty As Boolean
            Get
                Return GetValue(IsDirtyProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsDirtyProperty, value)
            End Set
        End Property

        Public Overloads Property Focusable As Boolean
            Get
                Return GetValue(FocusableProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(FocusableProperty, value)
            End Set
        End Property

#Region "Properties handled fully in XAML"
        Public Property ShowDropDownButton As Boolean
            Get
                Return GetValue(ShowDropDownButtonProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(ShowDropDownButtonProperty, value)
            End Set
        End Property

        Public Property Watermark As String
            Get
                Return GetValue(WatermarkProperty)
            End Get
            Set(value As String)
                SetCurrentValue(WatermarkProperty, value)
            End Set
        End Property

        Public Property MaxLength As Integer
            Get
                Return GetValue(MaxLengthProperty)
            End Get
            Set(value As Integer)
                SetCurrentValue(MaxLengthProperty, value)
            End Set
        End Property

        Public Property Icon As Object
            Get
                Return GetValue(IconProperty)
            End Get
            Set(value As Object)
                SetCurrentValue(IconProperty, value)
            End Set
        End Property

        Public Property IsReadOnly As Boolean
            Get
                Return GetValue(IsReadOnlyProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsReadOnlyProperty, value)
            End Set
        End Property

        Public Overloads Property IsTabStop As Boolean
            Get
                Return GetValue(IsTabStopProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsTabStopProperty, value)
            End Set
        End Property

        Public Overloads Property IsToolTipEnabled As Boolean
            Get
                Return GetValue(IsToolTipEnabledProperty)
            End Get
            Set(value As Boolean)
                SetCurrentValue(IsToolTipEnabledProperty, value)
            End Set
        End Property

        Public Overloads Property ItemTemplate As DataTemplate
            Get
                Return GetValue(ItemTemplateProperty)
            End Get
            Set(value As DataTemplate)
                SetCurrentValue(ItemTemplateProperty, value)
            End Set
        End Property

        Public Overloads Property LoadingContent As Object
            Get
                Return GetValue(LoadingContentProperty)
            End Get
            Set(value As Object)
                SetCurrentValue(LoadingContentProperty, value)
            End Set
        End Property
#End Region

    End Class
End Namespace