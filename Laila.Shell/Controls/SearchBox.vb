Imports System.Threading
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Input
Imports Laila.Shell.Helpers

Namespace Controls
    Public Class SearchBox
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(SearchBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly NavigationProperty As DependencyProperty = DependencyProperty.Register("Navigation", GetType(Navigation), GetType(SearchBox), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Private PART_TextBox As TextBox
        Private PART_CancelButton As Button
        Private _timer As Timer

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(SearchBox), New FrameworkPropertyMetadata(GetType(SearchBox)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            PART_TextBox = Me.Template.FindName("PART_TextBox", Me)
            PART_CancelButton = Me.Template.FindName("PART_CancelButton", Me)

            AddHandler Me.PART_TextBox.PreviewKeyUp,
                Sub(s As Object, e As KeyEventArgs)
                    If Not _timer Is Nothing Then
                        _timer.Dispose()
                    End If

                    _timer = New Timer(
                        Sub()
                            _timer.Dispose()
                            _timer = Nothing

                            UIHelper.OnUIThread(
                                Sub()
                                    If Not TypeOf Me.Folder Is SearchFolder Then
                                        Me.Folder = SearchFolder.FromTerms(Me.PART_TextBox.Text, Me.Folder)
                                    ElseIf Not String.IsNullOrWhiteSpace(Me.PART_TextBox.Text) Then
                                        CType(Me.Folder, SearchFolder).Update(Me.PART_TextBox.Text)
                                    Else
                                        Me.Folder.CancelEnumeration()
                                        Me.Navigation.Back()
                                    End If
                                End Sub)
                        End Sub, Nothing, 750, Timeout.Infinite)
                End Sub

            AddHandler Me.PART_CancelButton.Click,
                Sub(s As Object, e As EventArgs)
                    Me.PART_TextBox.Text = Nothing
                    Me.Folder.CancelEnumeration()
                    Me.Navigation.Back()
                End Sub
        End Sub

        Public Property Folder As Folder
            Get
                Return GetValue(FolderProperty)
            End Get
            Set(value As Folder)
                SetValue(FolderProperty, value)
            End Set
        End Property

        Public Property Navigation As Navigation
            Get
                Return GetValue(NavigationProperty)
            End Get
            Set(value As Navigation)
                SetValue(NavigationProperty, value)
            End Set
        End Property

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim sb As SearchBox = d
            If Not e.NewValue Is Nothing AndAlso TypeOf e.NewValue Is SearchFolder Then
                sb.PART_TextBox.Text = CType(e.NewValue, SearchFolder).Terms
            Else
                sb.PART_TextBox.Text = Nothing
            End If
        End Sub
    End Class
End Namespace