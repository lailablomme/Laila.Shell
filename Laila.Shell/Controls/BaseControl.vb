Imports System.ComponentModel
Imports System.Windows
Imports System.Windows.Controls
Imports System.Windows.Media
Imports Laila.Shell.Converters
Imports Laila.Shell.Themes

Namespace Controls
    Public MustInherit Class BaseControl
        Inherits Control

        Public Shared ReadOnly FolderProperty As DependencyProperty = DependencyProperty.Register("Folder", GetType(Folder), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnFolderChanged))
        Public Shared ReadOnly SelectedItemsProperty As DependencyProperty = DependencyProperty.Register("SelectedItems", GetType(IEnumerable(Of Item)), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnSelectedItemsChanged))
        Public Shared ReadOnly DoShowCheckBoxesToSelectProperty As DependencyProperty = DependencyProperty.Register("DoShowCheckBoxesToSelect", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowCheckBoxesToSelectOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowCheckBoxesToSelectOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowCheckBoxesToSelectOverrideChanged))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColor", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowEncryptedOrCompressedFilesInColorOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowEncryptedOrCompressedFilesInColorOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowEncryptedOrCompressedFilesInColorOverrideChanged))
        Public Shared ReadOnly IsDoubleClickToOpenItemProperty As DependencyProperty = DependencyProperty.Register("IsDoubleClickToOpenItem", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsDoubleClickToOpenItemOverrideProperty As DependencyProperty = DependencyProperty.Register("IsDoubleClickToOpenItemOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsDoubleClickToOpenItemOverrideChanged))
        Public Shared ReadOnly IsUnderlineItemOnHoverProperty As DependencyProperty = DependencyProperty.Register("IsUnderlineItemOnHover", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsUnderlineItemOnHoverOverrideProperty As DependencyProperty = DependencyProperty.Register("IsUnderlineItemOnHoverOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsUnderlineItemOnHoverOverrideChanged))
        Public Shared ReadOnly DoShowIconsOnlyProperty As DependencyProperty = DependencyProperty.Register("DoShowIconsOnly", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowIconsOnlyOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowIconsOnlyOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowIconsOnlyOverrideChanged))
        Public Shared ReadOnly DoShowTypeOverlayProperty As DependencyProperty = DependencyProperty.Register("DoShowTypeOverlay", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowTypeOverlayOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowTypeOverlayOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowTypeOverlayOverrideChanged))
        Public Shared ReadOnly DoShowFolderContentsInInfoTipProperty As DependencyProperty = DependencyProperty.Register("DoShowFolderContentsInInfoTip", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowFolderContentsInInfoTipOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowFolderContentsInInfoTipOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowFolderContentsInInfoTipOverrideChanged))
        Public Shared ReadOnly DoShowInfoTipsProperty As DependencyProperty = DependencyProperty.Register("DoShowInfoTips", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowInfoTipsOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowInfoTipsOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowInfoTipsOverrideChanged))
        Public Shared ReadOnly IsCompactModeProperty As DependencyProperty = DependencyProperty.Register("IsCompactMode", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly IsCompactModeOverrideProperty As DependencyProperty = DependencyProperty.Register("IsCompactModeOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnIsCompactModeOverrideChanged))
        Public Shared ReadOnly DoTypeToSelectProperty As DependencyProperty = DependencyProperty.Register("DoTypeToSelect", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoTypeToSelectOverrideProperty As DependencyProperty = DependencyProperty.Register("DoTypeToSelectOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoTypeToSelectOverrideChanged))
        Public Shared ReadOnly DoUseWindows11ExplorerMenuProperty As DependencyProperty = DependencyProperty.Register("DoUseWindows11ExplorerMenu", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoUseWindows11ExplorerMenuOverrideProperty As DependencyProperty = DependencyProperty.Register("DoUseWindows11ExplorerMenuOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoUseWindows11ExplorerMenuOverrideChanged))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeView", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAllFoldersInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAllFoldersInTreeViewOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAllFoldersInTreeViewOverrideChanged))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeView", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowAvailabilityStatusInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowAvailabilityStatusInTreeViewOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowAvailabilityStatusInTreeViewOverrideChanged))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolder", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoExpandTreeViewToCurrentFolderOverrideProperty As DependencyProperty = DependencyProperty.Register("DoExpandTreeViewToCurrentFolderOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoExpandTreeViewToCurrentFolderOverrideChanged))
        Public Shared ReadOnly DoShowLibrariesInTreeViewProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeView", GetType(Boolean), GetType(BaseControl), New FrameworkPropertyMetadata(False, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))
        Public Shared ReadOnly DoShowLibrariesInTreeViewOverrideProperty As DependencyProperty = DependencyProperty.Register("DoShowLibrariesInTreeViewOverride", GetType(Boolean?), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault, AddressOf OnDoShowLibrariesInTreeViewOverrideChanged))
        Public Shared ReadOnly ColorsProperty As DependencyProperty = DependencyProperty.Register("Colors", GetType(StandardColors), GetType(BaseControl), New FrameworkPropertyMetadata(Nothing, FrameworkPropertyMetadataOptions.BindsTwoWayByDefault))

        Shared Sub New()
            DefaultStyleKeyProperty.OverrideMetadata(GetType(BaseControl), New FrameworkPropertyMetadata(GetType(BaseControl)))
        End Sub

        Public Overrides Sub OnApplyTemplate()
            MyBase.OnApplyTemplate()

            If Me.Colors Is Nothing Then Me.Colors = New StandardColors()

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
                        Case "DoShowEncryptedOrCompressedFilesInColor"
                            setDoShowEncryptedOrCompressedFilesInColor()
                        Case "IsCompactMode"
                            setIsCompactMode()
                        Case "DoShowAllFoldersInTreeView"
                            setDoShowAllFoldersInTreeView()
                        Case "DoShowAvailabilityStatusInTreeView"
                            setDoShowAvailabilityStatusInTreeView()
                        Case "DoExpandTreeViewToCurrentFolder"
                            setDoExpandTreeViewToCurrentFolder()
                        Case "DoShowLibrariesInTreeView"
                            setDoShowLibrariesInTreeView()
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
            setDoUseWindows11ExplorerMenu()
            setDoShowEncryptedOrCompressedFilesInColor()
            setIsCompactMode()
            setDoShowAllFoldersInTreeView()
            setDoShowAvailabilityStatusInTreeView()
            setDoExpandTreeViewToCurrentFolder()
            setDoShowLibrariesInTreeView()
            setDoUseWindows11ExplorerMenu()
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
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
            Dim bfv As BaseControl = d
            bfv.setDoTypeToSelect()
        End Sub

        Public Property DoUseWindows11ExplorerMenu As Boolean
            Get
                Return GetValue(DoUseWindows11ExplorerMenuProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoUseWindows11ExplorerMenuProperty, value)
            End Set
        End Property

        Private Sub setDoUseWindows11ExplorerMenu()
            If Me.DoUseWindows11ExplorerMenuOverride.HasValue Then
                Me.DoUseWindows11ExplorerMenu = Me.DoUseWindows11ExplorerMenuOverride.Value
            Else
                Me.DoUseWindows11ExplorerMenu = Shell.Settings.DoUseWindows11ExplorerMenu
            End If
        End Sub

        Public Property DoUseWindows11ExplorerMenuOverride As Boolean?
            Get
                Return GetValue(DoUseWindows11ExplorerMenuOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoUseWindows11ExplorerMenuOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoUseWindows11ExplorerMenuOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseControl = d
            bfv.setDoUseWindows11ExplorerMenu()
        End Sub

        Public Property DoShowAllFoldersInTreeView As Boolean
            Get
                Return GetValue(DoShowAllFoldersInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAllFoldersInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAllFoldersInTreeView()
            If Me.DoShowAllFoldersInTreeViewOverride.HasValue Then
                Me.DoShowAllFoldersInTreeView = Me.DoShowAllFoldersInTreeViewOverride.Value
            Else
                Me.DoShowAllFoldersInTreeView = Shell.Settings.DoShowAllFoldersInTreeView
            End If
        End Sub

        Public Property DoShowAllFoldersInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAllFoldersInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAllFoldersInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAllFoldersInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseControl = d
            bfv.setDoShowAllFoldersInTreeView()
        End Sub

        Public Property DoShowAvailabilityStatusInTreeView As Boolean
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowAvailabilityStatusInTreeView()
            If Me.DoShowAvailabilityStatusInTreeViewOverride.HasValue Then
                Me.DoShowAvailabilityStatusInTreeView = Me.DoShowAvailabilityStatusInTreeViewOverride.Value
            Else
                Me.DoShowAvailabilityStatusInTreeView = Shell.Settings.DoShowAvailabilityStatusInTreeView
            End If
        End Sub

        Public Property DoShowAvailabilityStatusInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowAvailabilityStatusInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowAvailabilityStatusInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowAvailabilityStatusInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseControl = d
            bfv.setDoShowAvailabilityStatusInTreeView()
        End Sub

        Public Property DoExpandTreeViewToCurrentFolder As Boolean
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderProperty, value)
            End Set
        End Property

        Private Sub setDoExpandTreeViewToCurrentFolder()
            If Me.DoExpandTreeViewToCurrentFolderOverride.HasValue Then
                Me.DoExpandTreeViewToCurrentFolder = Me.DoExpandTreeViewToCurrentFolderOverride.Value
            Else
                Me.DoExpandTreeViewToCurrentFolder = Shell.Settings.DoExpandTreeViewToCurrentFolder
            End If
        End Sub

        Public Property DoExpandTreeViewToCurrentFolderOverride As Boolean?
            Get
                Return GetValue(DoExpandTreeViewToCurrentFolderOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoExpandTreeViewToCurrentFolderOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoExpandTreeViewToCurrentFolderOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseControl = d
            bfv.setDoExpandTreeViewToCurrentFolder()
        End Sub

        Public Property DoShowLibrariesInTreeView As Boolean
            Get
                Return GetValue(DoShowLibrariesInTreeViewProperty)
            End Get
            Protected Set(ByVal value As Boolean)
                SetCurrentValue(DoShowLibrariesInTreeViewProperty, value)
            End Set
        End Property

        Private Sub setDoShowLibrariesInTreeView()
            If Me.DoShowLibrariesInTreeViewOverride.HasValue Then
                Me.DoShowLibrariesInTreeView = Me.DoShowLibrariesInTreeViewOverride.Value
            Else
                Me.DoShowLibrariesInTreeView = Shell.Settings.DoShowLibrariesInTreeView
            End If
        End Sub

        Public Property DoShowLibrariesInTreeViewOverride As Boolean?
            Get
                Return GetValue(DoShowLibrariesInTreeViewOverrideProperty)
            End Get
            Set(ByVal value As Boolean?)
                SetCurrentValue(DoShowLibrariesInTreeViewOverrideProperty, value)
            End Set
        End Property

        Public Shared Sub OnDoShowLibrariesInTreeViewOverrideChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bfv As BaseControl = d
            bfv.setDoShowLibrariesInTreeView()
        End Sub

        Public Property Colors As StandardColors
            Get
                Return GetValue(ColorsProperty)
            End Get
            Protected Set(ByVal value As StandardColors)
                SetCurrentValue(ColorsProperty, value)
            End Set
        End Property

        Protected Overridable Sub OnFolderChanged(ByVal e As DependencyPropertyChangedEventArgs)

        End Sub

        Protected Overridable Sub OnSelectedItemsChanged(ByVal e As DependencyPropertyChangedEventArgs)

        End Sub

        Shared Sub OnFolderChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bc As BaseControl = d
            bc.OnFolderChanged(e)
        End Sub

        Shared Sub OnSelectedItemsChanged(ByVal d As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim bc As BaseControl = d
            bc.OnSelectedItemsChanged(e)
        End Sub
    End Class
End Namespace