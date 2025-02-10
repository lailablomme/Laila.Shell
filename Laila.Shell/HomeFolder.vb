Imports System.Drawing
Imports System.Runtime.InteropServices
Imports System.Threading
Imports System.Windows
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class HomeFolder
    Inherits Folder

    Public Sub New(parent As Folder, doKeepAlive As Boolean)
        MyBase.New(Nothing, parent, doKeepAlive, False, Nothing)

        _pidl = Pidl.FromCustomBytes(System.Text.UTF8Encoding.Default.GetBytes(Me.FullPath))
        Me.HasSubFolders = False
        _columns = New List(Of Column)() From {
            New Column(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:10"), New CM_COLUMNINFO(), 0) With {.IsVisible = True},
            New Column(New PROPERTYKEY("B725F130-47EF-101A-A5F1-02608C9EEBAC:16"), New CM_COLUMNINFO(), 0) With {.IsVisible = True},
            New Column(New PROPERTYKEY("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6"), New CM_COLUMNINFO(), 0) With {.IsVisible = True}
        }
        Me.ItemsGroupByPropertyName = "PropertiesByKeyAsText[" & Home_CategoryProperty.Key.ToString() & "].GroupByText"
        Me.ItemsSortDirection = ComponentModel.ListSortDirection.Descending
    End Sub

    Public Overrides ReadOnly Property DisplayName As String
        Get
            Return "Home"
        End Get
    End Property

    Public Overrides ReadOnly Property FullPath As String
        Get
            Return "::{b8b10b36-5c36-4f45-ae9a-79f0297d64e1}"
        End Get
    End Property

    Public Overrides ReadOnly Property Attributes As SFGAO
        Get
            Return SFGAO.FOLDER
        End Get
    End Property

    Public Overrides ReadOnly Property Parent As Folder
        Get
            Return Shell.Desktop
        End Get
    End Property

    Friend Overrides Function GetShellFolderOnCurrentThread() As IShellFolderForIContextMenu
        Return Nothing
    End Function

    Public Overrides ReadOnly Property Icon(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr, result As ImageSource, shellItem As IShellItem2
            Try
                shellItem = Item.GetIShellItem2FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
                Dim h As HRESULT = CType(shellItem, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), SIIGBF.SIIGBF_ICONONLY, hbitmap)
                If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                    result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Else
                    Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                End If
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
                If Not shellItem Is Nothing Then
                    Marshal.ReleaseComObject(shellItem)
                    shellItem = Nothing
                End If
            End Try
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property Image(size As Integer) As ImageSource
        Get
            Dim hbitmap As IntPtr, result As ImageSource, shellItem As IShellItem2
            Try
                shellItem = Item.GetIShellItem2FromParsingName("shell:::{679f85cb-0220-4080-b29b-5540cc05aab6}")
                Dim h As HRESULT = CType(shellItem, IShellItemImageFactory).GetImage(New System.Drawing.Size(size * Settings.DpiScaleX, size * Settings.DpiScaleY), 0, hbitmap)
                If h = 0 AndAlso Not IntPtr.Zero.Equals(hbitmap) Then
                    result = Interop.Imaging.CreateBitmapSourceFromHBitmap(hbitmap, IntPtr.Zero, Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions())
                Else
                    Debug.WriteLine("IShellItemImageFactory.GetImage failed with hresult " & h.ToString())
                End If
            Finally
                If Not IntPtr.Zero.Equals(hbitmap) Then
                    Functions.DeleteObject(hbitmap)
                    hbitmap = IntPtr.Zero
                End If
                If Not shellItem Is Nothing Then
                    Marshal.ReleaseComObject(shellItem)
                    shellItem = Nothing
                End If
            End Try
            Return result
        End Get
    End Property

    Public Overrides ReadOnly Property OverlayImage(size As Integer) As ImageSource
        Get
            Return Nothing
        End Get
    End Property

    Protected Overrides Sub EnumerateItems(shellItem2 As IShellItem2, flags As UInteger, cancellationToken As CancellationToken,
        isSortPropertyByText As Boolean, isSortPropertyDisplaySortValue As Boolean, sortPropertyKey As String, result As Dictionary(Of String, Item), newFullPaths As HashSet(Of String), addItems As Action)

        Me.HasSubFolders = False

        Dim count As UInt64 = UInt64.MaxValue
        For Each item In PinnedItems.GetPinnedItems()
            Dim clone As Item = item.Clone()
            clone.LogicalParent = Me
            clone.TreeSortPrefix = String.Format("{0:00000000000000000000}", count)
            clone.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)

            Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.PINNED_ITEM)
            clone._propertiesByKey.Add(categoryProperty.Key.ToString(), categoryProperty)

            result.Add(clone.FullPath & "_" & clone.DisplayName, clone)
            newFullPaths.Add(clone.FullPath & "_" & clone.DisplayName)
            If TypeOf item Is Folder Then Me.HasSubFolders = True

            count -= 1
        Next

        For Each item In FrequentFolders.GetMostFrequent()
            Dim clone As Item = item.Clone()
            clone.LogicalParent = Me
            clone.TreeSortPrefix = String.Format("{0:00000000000000000000}", count)
            clone.ItemNameDisplaySortValuePrefix = String.Format("{0:00000000000000000000}", count)

            Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.FREQUENT_FOLDER)
            clone._propertiesByKey.Add(categoryProperty.Key.ToString(), categoryProperty)

            result.Add(clone.FullPath & "_" & clone.DisplayName, clone)
            newFullPaths.Add(clone.FullPath & "_" & clone.DisplayName)
            If TypeOf item Is Folder Then Me.HasSubFolders = True

            count -= 1
        Next

        MyBase.EnumerateItems(Shell.GetSpecialFolder(SpecialFolders.Recent).ShellItem2, flags, cancellationToken, isSortPropertyByText,
            isSortPropertyDisplaySortValue, sortPropertyKey, result, newFullPaths, addItems)
    End Sub

    Protected Overrides Function InitializeItem(item As Item) As Item
        MyBase.InitializeItem(item)

        If TypeOf item Is Link Then
            Dim link As Link = item
            Dim target As Item = Item.FromPidl(link.TargetPidl.AbsolutePIDL, Me, False, True)
            If Not TypeOf target Is Folder _
                AndAlso Not String.IsNullOrWhiteSpace(target.PropertiesByKeyAsText("E3E0584C-B788-4A5A-BB20-7F5A44C9ACDD:6").Value) Then
                Dim modifiedProperty As [Property] = item.PropertiesByKeyAsText("b725f130-47ef-101a-a5f1-02608c9eebac:14")
                target.ItemNameDisplaySortValuePrefix = String.Format("{0:yyyyMMddHHmmssffff}", modifiedProperty.Value)

                Dim lastAccessedProperty As [Property] = target.PropertiesByKeyAsText("B725F130-47EF-101A-A5F1-02608C9EEBAC:16")
                lastAccessedProperty._rawValue = modifiedProperty._rawValue

                Dim categoryProperty As Home_CategoryProperty = New Home_CategoryProperty(Home_CategoryProperty.Type.RECENT_FILE)
                target._propertiesByKey.Add(categoryProperty.Key.ToString(), categoryProperty)

                item.Dispose()
                Return target
            Else
                target.Dispose()
            End If
        End If

        item.Dispose()
        Return Nothing
    End Function

    Public Overrides ReadOnly Property CanSort As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides ReadOnly Property CanGroupBy As Boolean
        Get
            Return False
        End Get
    End Property

    Public Overrides Function Clone() As Item
        Return Shell.RunOnSTAThread(
            Function() As Item
                Return New HomeFolder(Nothing, _doKeepAlive)
            End Function)
    End Function
End Class
