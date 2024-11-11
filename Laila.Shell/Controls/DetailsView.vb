Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Helpers
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.ComponentModel
Imports System.Windows.Data
Imports System.Windows.Media

Namespace Controls
    Public Class DetailsView
        Inherits BaseFolderView

        Private PART_Ext As Behaviors.GridViewExtBehavior
        Private _isLoaded As Boolean

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(DetailsView), New FrameworkPropertyMetadata(GetType(DetailsView)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            AddHandler PART_ListView.Loaded,
                Sub(s As Object, e As EventArgs)
                    If Not _isLoaded Then
                        _isLoaded = True

                        ' notify of sort/group by changes
                        Me.PART_Ext = Microsoft.Xaml.Behaviors.Interaction.GetBehaviors(Me.PART_ListView).FirstOrDefault(Function(b) TypeOf b Is Behaviors.GridViewExtBehavior)
                        If Not Me.PART_Ext Is Nothing Then
                            AddHandler Me.PART_Ext.SortChanged,
                                Sub(s2 As Object, e2 As EventArgs)
                                    Me.Folder.NotifyOfPropertyChange("ItemsSortPropertyName")
                                End Sub
                            AddHandler Me.PART_Ext.GroupByChanged,
                                Sub(s2 As Object, e2 As EventArgs)
                                    Me.Folder.NotifyOfPropertyChange("ItemsGroupByPropertyName")
                                End Sub
                        End If
                    End If
                End Sub
        End Sub

        Public Function buildColumnsIn() As Behaviors.GridViewExtBehavior.ColumnsInData
            Dim d As Behaviors.GridViewExtBehavior.ColumnsInData = New Behaviors.GridViewExtBehavior.ColumnsInData()
            d.ViewName = Me.Folder.FullPath
            d.PrimarySortProperties = "PrimarySort"
            d.Items = New List(Of GridViewColumn)()

            For Each column In Me.Folder.Columns.Where(Function(c) Not String.IsNullOrWhiteSpace(c.DisplayName))
                Dim [property] As [Property] = [Property].FromCanonicalName(column.CanonicalName)

                Dim gvc As GridViewColumn = New GridViewColumn()
                gvc.Header = column.DisplayName
                gvc.CellTemplate = getCellTemplate(column, [property])
                gvc.SetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty, column.IsVisible)
                gvc.SetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                If column.CanonicalName = "System.ItemNameDisplay" Then
                    gvc.SetValue(Behaviors.GridViewExtBehavior.SortPropertyNameProperty, "ItemNameDisplaySortValue")
                    gvc.SetValue(Behaviors.GridViewExtBehavior.CanHideProperty, False)
                End If
                gvc.SetValue(Behaviors.GridViewExtBehavior.GroupByPropertyNameProperty, String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString()))

                Dim isVisibleDescriptor As DependencyPropertyDescriptor =
                    DependencyPropertyDescriptor.FromProperty(Behaviors.GridViewExtBehavior.IsVisibleProperty, gvc.GetType())
                isVisibleDescriptor.AddValueChanged(gvc,
                    Sub(s2 As Object, e2 As EventArgs)
                        column.IsVisible = gvc.GetValue(Behaviors.GridViewExtBehavior.IsVisibleProperty)
                    End Sub)

                d.Items.Add(gvc)
            Next

            Return d
        End Function

        Private Function getCellTemplate(column As Column, [property] As [Property]) As DataTemplate
            Dim template As DataTemplate = New DataTemplate()

            Dim gridFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(Grid))
            Dim columnDefinition1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition1.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition1)
            Dim columnDefinition2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition2.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Auto))
            gridFactory.AppendChild(columnDefinition2)
            Dim columnDefinition3 As FrameworkElementFactory = New FrameworkElementFactory(GetType(ColumnDefinition))
            columnDefinition3.SetValue(ColumnDefinition.WidthProperty, New GridLength(1, GridUnitType.Star))
            gridFactory.AppendChild(columnDefinition3)

            Dim getIconFactory As Func(Of String, FrameworkElementFactory) =
                Function(bindTo As String) As FrameworkElementFactory
                    Dim imageFactory1 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                    imageFactory1.SetValue(Grid.ColumnProperty, 1)
                    imageFactory1.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                    imageFactory1.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                    imageFactory1.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                    imageFactory1.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                    imageFactory1.SetValue(Image.SourceProperty, New Binding() With {
                        .Path = New PropertyPath(bindTo),
                        .Mode = BindingMode.OneWay,
                        .IsAsync = True
                    })
                    Dim style As Style = New Style(GetType(Image))
                    style.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Collapsed))
                    style.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                    Dim dataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding(String.Format("ColumnIndexFor[PropertiesByKeyAsText[{0}].Value]", column.PROPERTYKEY.ToString())) With
                                   {
                                       .ElementName = "ext",
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = 0
                    }
                    dataTrigger1.Setters.Add(New Setter(Image.VisibilityProperty, Visibility.Visible))
                    style.Triggers.Add(dataTrigger1)
                    Dim dataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsHidden") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger2.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger2)
                    Dim dataTrigger3 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                    dataTrigger3.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                    style.Triggers.Add(dataTrigger3)
                    imageFactory1.SetValue(Image.StyleProperty, style)
                    Return imageFactory1
                End Function
            gridFactory.AppendChild(getIconFactory("IconAsync[16]"))
            gridFactory.AppendChild(getIconFactory("OverlaySmallAsync"))

            If [property].HasIcon Then
                Dim imageFactory2 As FrameworkElementFactory = New FrameworkElementFactory(GetType(Image))
                imageFactory2.SetValue(Grid.ColumnProperty, 1)
                imageFactory2.SetValue(Image.MarginProperty, New Thickness(0, 0, 4, 0))
                imageFactory2.SetValue(Image.WidthProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HeightProperty, Convert.ToDouble(16))
                imageFactory2.SetValue(Image.HorizontalAlignmentProperty, HorizontalAlignment.Left)
                imageFactory2.SetValue(Image.VerticalAlignmentProperty, VerticalAlignment.Center)
                imageFactory2.SetValue(Image.SourceProperty, New Binding() With {
                    .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Icon16Async", column.PROPERTYKEY.ToString())),
                    .Mode = BindingMode.OneWay,
                    .IsAsync = True
                })
                Dim imageStyle As Style = New Style(GetType(Image))
                imageStyle.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(1)))
                Dim imageDataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
                imageDataTrigger1.Setters.Add(New Setter(Image.OpacityProperty, Convert.ToDouble(0.5)))
                imageStyle.Triggers.Add(imageDataTrigger1)
                imageFactory2.SetValue(Image.StyleProperty, imageStyle)
                gridFactory.AppendChild(imageFactory2)
            End If

            Dim textBlockFactory As FrameworkElementFactory = New FrameworkElementFactory(GetType(TextBlock))
            textBlockFactory.SetValue(Grid.ColumnProperty, 2)
            textBlockFactory.SetValue(TextBlock.TextAlignmentProperty, column.Alignment)
            textBlockFactory.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center)
            textBlockFactory.SetValue(TextBlock.TextTrimmingProperty, TextTrimming.CharacterEllipsis)
            textBlockFactory.SetValue(TextBlock.PaddingProperty, New Thickness(0, 0, 2, 0))
            textBlockFactory.SetValue(TextBlock.TextProperty, New Binding() With {
                .Path = New PropertyPath(String.Format("PropertiesByKeyAsText[{0}].Text", column.PROPERTYKEY.ToString())),
                .Mode = BindingMode.OneWay
            })
            Dim textBlockStyle As Style = New Style(GetType(TextBlock))
            textBlockStyle.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Black))
            textBlockStyle.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(1)))
            Dim textBlockDataTrigger1 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCut") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger1.Setters.Add(New Setter(TextBlock.OpacityProperty, Convert.ToDouble(0.5)))
            textBlockStyle.Triggers.Add(textBlockDataTrigger1)
            Dim textBlockDataTrigger2 As DataTrigger = New DataTrigger() With {
                        .Binding = New Binding("IsCompressed") With
                                   {
                                       .Mode = BindingMode.OneWay
                                   },
                        .Value = True
                    }
            textBlockDataTrigger2.Setters.Add(New Setter(TextBlock.ForegroundProperty, Brushes.Blue))
            textBlockStyle.Triggers.Add(textBlockDataTrigger2)
            textBlockFactory.SetValue(Image.StyleProperty, textBlockStyle)
            gridFactory.AppendChild(textBlockFactory)

            template.VisualTree = gridFactory
            Return template
        End Function

        Protected Overrides Sub ClearBinding()
            MyBase.ClearBinding()

            If Not Me.PART_ListView Is Nothing Then
                CType(Me.PART_ListView.View, GridView).Columns.Clear()
            End If
        End Sub

        Protected Overrides Sub MakeBinding(folder As Folder)
            MyBase.MakeBinding(folder)

            If Not Me.PART_ListView Is Nothing Then
                Me.ColumnsIn = buildColumnsIn()
            End If
        End Sub

        Protected Overrides Sub Folder_PropertyChanged(s As Object, e As PropertyChangedEventArgs)
            Select Case e.PropertyName
                Case "ItemsSortPropertyName", "ItemsSortDirection"
                    If Not Me.PART_Ext Is Nothing Then
                        Me.PART_Ext.UpdateSortGlyphs()
                    End If
            End Select

            MyBase.Folder_PropertyChanged(s, e)
        End Sub

        Protected Overrides Sub GetItemNameCoordinates(listViewItem As ListViewItem, ByRef textAlignment As TextAlignment,
                                                       ByRef point As Point, ByRef size As Size, ByRef fontSize As Double)
            Dim column As Column = Me.Folder.Columns("System.ItemNameDisplay")
            If Not column Is Nothing Then
                Dim headers As IEnumerable(Of GridViewColumnHeader) =
                                            UIHelper.FindVisualChildren(Of GridViewColumnHeader)(Me.PART_ListView)
                Dim header As GridViewColumnHeader =
                                            headers.FirstOrDefault(Function(h) Not h.Column Is Nothing _
                                                AndAlso h.Column.GetValue(Behaviors.GridViewExtBehavior.PropertyNameProperty) _
                                                    = String.Format("PropertiesByKeyAsText[{0}].Value", column.PROPERTYKEY.ToString()))
                If Not header Is Nothing Then
                    Dim width As Double = header.ActualWidth
                    Dim ptLeft As Point = Me.PointFromScreen(header.PointToScreen(New Point(0, 0)))
                    If header.Column.GetValue(Behaviors.GridViewExtBehavior.ColumnIndexProperty) = 0 Then
                        ptLeft.X += 20
                        width -= 20
                    End If
                    Dim ptTop As Point = Me.PointFromScreen(listViewItem.PointToScreen(New Point(0, 0)))

                    point.X = ptLeft.X + 5
                    point.Y = ptTop.Y
                    size.Width = width - 5
                    size.Height = listViewItem.ActualHeight
                End If
            End If
            textAlignment = TextAlignment.Left
            fontSize = Me.FontSize
        End Sub

        Public Property ColumnsIn As Behaviors.GridViewExtBehavior.ColumnsInData
            Get
                Return GetValue(ColumnsInProperty)
            End Get
            Set(ByVal value As Behaviors.GridViewExtBehavior.ColumnsInData)
                SetCurrentValue(ColumnsInProperty, value)
            End Set
        End Property
    End Class
End Namespace