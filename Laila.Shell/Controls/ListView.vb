Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Data
Imports WpfToolkit.Controls

Namespace Controls
    Public Class ListView
        Inherits BaseFolderView

        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(ListView), New FrameworkPropertyMetadata(GetType(ListView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            AddHandler PART_ListView.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        If Not Me.Folder Is Nothing Then setWrapPanel(Me.Folder)
                    End If
                End Sub
        End Sub

        Protected Overrides Sub Folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            MyBase.Folder_PropertyChanged(s, e)

            Select Case e.PropertyName
                Case "ItemsGroupByPropertyName"
                    setWrapPanel(s)
            End Select
        End Sub

        Protected Overrides Sub MakeBinding(folder As Folder)
            MyBase.MakeBinding(folder)

            setWrapPanel(folder)
        End Sub

        Private Sub setWrapPanel(folder As Folder)
            If Not Me.PART_ListView Is Nothing Then
                Dim wrapPanelFactory As FrameworkElementFactory
                If folder.ItemsGroupByPropertyName Is Nothing Then
                    wrapPanelFactory = New FrameworkElementFactory(GetType(VirtualizingWrapPanel))
                    wrapPanelFactory.SetValue(VirtualizingWrapPanel.OrientationProperty, Orientation.Vertical)
                    wrapPanelFactory.SetValue(VirtualizingWrapPanel.MarginProperty, New Thickness(20, 0, -25, 0))
                    wrapPanelFactory.SetValue(VirtualizingWrapPanel.AllowDifferentSizedItemsProperty, True)
                    wrapPanelFactory.SetValue(VirtualizingWrapPanel.SpacingModeProperty, SpacingMode.None)
                Else
                    wrapPanelFactory = New FrameworkElementFactory(GetType(WrapPanel))
                    wrapPanelFactory.SetValue(WrapPanel.OrientationProperty, Orientation.Vertical)
                    wrapPanelFactory.SetValue(WrapPanel.MarginProperty, New Thickness(20, 0, -25, 0))
                End If

                Dim itemsPanelTemplate As New ItemsPanelTemplate()
                itemsPanelTemplate.VisualTree = wrapPanelFactory
                Me.PART_ListView.ItemsPanel = itemsPanelTemplate
            End If
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))
            point.X += 16 + 4 + 2 + 16 + 4
            point.Y += 1
            listViewItem.Measure(New Size(Double.PositiveInfinity, Double.PositiveInfinity))
            size.Width = listViewItem.DesiredSize.Width - 16 - 4 - 2
            size.Height = listViewItem.DesiredSize.Height
            textAlignment = TextAlignment.Left
            fontSize = Me.FontSize
        End Sub
    End Class
End Namespace