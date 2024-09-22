Imports System.Collections.ObjectModel
Imports System.Threading
Imports Laila.Shell.Data
Imports System.Runtime.InteropServices
Imports System.Windows
Imports System.ComponentModel
Imports System.Windows.Data

Public Class TreeViewFolder
    Inherits Folder
    Implements ITreeViewItemData

    Private _isSelected As Boolean
    Private _isExpanded As Boolean
    Private _parent As ITreeViewItemData
    Friend _folders As ObservableCollection(Of TreeViewFolder)
    Private _isLoading As Boolean
    Private _fromThread As Boolean

    Public Overloads Shared Function FromParsingName(parsingName As String, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean)) As TreeViewFolder
        Dim shellItem2 As IShellItem2 = GetIShellItem2FromParsingName(parsingName)
        Dim attr As Integer = SFGAO.FOLDER
        shellItem2.GetAttributes(attr, attr)
        If CBool(attr And SFGAO.FOLDER) Then
            Return New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(shellItem2), shellItem2, logicalParent, setIsLoadingAction)
        Else
            Throw New InvalidOperationException("Only folders.")
        End If
    End Function

    Public Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(shellFolder, shellItem2, logicalParent, setIsLoadingAction)
    End Sub

    Public Sub New(bindingParent As Folder, pidl As IntPtr, logicalParent As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(bindingParent, pidl, logicalParent, setIsLoadingAction)
    End Sub

    Public Property IsSelected As Boolean Implements ITreeViewItemData.IsSelected
        Get
            Return _isSelected
        End Get
        Set(value As Boolean)
            SetValue(_isSelected, value)
        End Set
    End Property

    Public Property IsExpanded As Boolean Implements ITreeViewItemData.IsExpanded
        Get
            Return _isExpanded
        End Get
        Set(value As Boolean)
            SetValue(_isExpanded, value)
            If value AndAlso (_folders Is Nothing OrElse (_folders.Count = 1 AndAlso TypeOf _folders(0) Is DummyTreeViewFolder)) Then
                _folders = Nothing
                NotifyOfPropertyChange("FoldersThreaded")
            End If
        End Set
    End Property

    Private Property ITreeViewItemData_Parent As ITreeViewItemData Implements ITreeViewItemData.Parent
        Get
            Return _parent
        End Get
        Set(value As ITreeViewItemData)
            SetValue(_parent, value)
        End Set
    End Property

    Public Property IsLoading As Boolean
        Get
            Return _isLoading
        End Get
        Set(value As Boolean)
            SetValue(_isLoading, value)
        End Set
    End Property

    Public ReadOnly Property FoldersThreaded As ObservableCollection(Of TreeViewFolder)
        Get
            If _folders Is Nothing AndAlso (Me.IsExpanded OrElse Me.ITreeViewItemData_Parent Is Nothing OrElse Me.ITreeViewItemData_Parent.IsExpanded) Then
                If Not _setIsLoadingAction Is Nothing Then
                    _setIsLoadingAction(True)
                End If

                Dim t As Thread = New Thread(New ThreadStart(
                    Sub()
                        _fromThread = True
                        Dim result As ObservableCollection(Of TreeViewFolder) = Me.Folders
                        _fromThread = False

                        If Not _setIsLoadingAction Is Nothing Then
                            _setIsLoadingAction(False)
                        End If
                    End Sub))

                t.Start()

                Return New ObservableCollection(Of TreeViewFolder)() From {
                    New DummyTreeViewFolder("Loading...")
                }
            Else
                Return _folders
            End If
        End Get
    End Property

    Public Overridable Property Folders As ObservableCollection(Of TreeViewFolder)
        Get
            If _folders Is Nothing Then
                Dim result As ObservableCollection(Of TreeViewFolder) = New ObservableCollection(Of TreeViewFolder)()

                If Not isWindows7OrLower() Then
                    Dim bindCtx As ComTypes.IBindCtx, bindCtxPtr As IntPtr
                    Functions.CreateBindCtx(0, bindCtxPtr)
                    bindCtx = Marshal.GetTypedObjectForIUnknown(bindCtxPtr, GetType(ComTypes.IBindCtx))

                    Dim propertyBag As IPropertyBag, propertyBagPtr As IntPtr
                    Functions.PSCreateMemoryPropertyStore(GetType(IPropertyBag).GUID, propertyBagPtr)
                    propertyBag = Marshal.GetTypedObjectForIUnknown(propertyBagPtr, GetType(IPropertyBag))

                    Dim var As New PROPVARIANT()
                    var.vt = VarEnum.VT_UI4
                    var.union.uintVal = CType(SHCONTF.FOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, UInt32) 'Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN Or SHCONTF.FOLDERS
                    propertyBag.Write("SHCONTF", var)

                    bindCtx.RegisterObjectParam("SHBindCtxPropertyBag", propertyBag)
                    Dim ptr2 As IntPtr
                    bindCtxPtr = Marshal.GetIUnknownForObject(bindCtx)

                    ShellItem2.BindToHandler(bindCtxPtr, Guids.BHID_EnumItems, GetType(IEnumShellItems).GUID, ptr2)
                    Dim enumShellItems As IEnumShellItems = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IEnumShellItems))

                    Try
                        Dim shellItemArray(0) As IShellItem, feteched As UInt32 = 1
                        Application.Current.Dispatcher.Invoke(
                            Sub()
                                enumShellItems.Next(1, shellItemArray, feteched)
                            End Sub)
                        Dim onlyOnce As Boolean = True
                        While feteched = 1 AndAlso (Me.IsExpanded OrElse onlyOnce OrElse Not _fromThread)
                            Dim tvf As TreeViewFolder = New TreeViewFolder(Folder.GetIShellFolderFromIShellItem2(shellItemArray(0)), shellItemArray(0), Me, _setIsLoadingAction)
                            tvf.ITreeViewItemData_Parent = Me
                            result.Add(tvf)
                            Application.Current.Dispatcher.Invoke(
                                Sub()
                                    enumShellItems.Next(1, shellItemArray, feteched)
                                End Sub)
                            onlyOnce = False
                            Thread.Sleep(10)
                        End While
                        If Not Me.IsExpanded AndAlso result.Count > 0 AndAlso _fromThread Then
                            result.Clear()
                            result.Add(New DummyTreeViewFolder("Loading..."))
                        End If
                    Catch ex As Exception
                    End Try
                Else
                    Dim list As IEnumIDList
                    Me.ShellFolder.EnumObjects(Nothing, SHCONTF.FOLDERS Or SHCONTF.INCLUDEHIDDEN Or SHCONTF.INCLUDESUPERHIDDEN, list)
                    If Not list Is Nothing Then
                        Dim pidl(0) As IntPtr, fetched As Integer
                        While list.Next(1, pidl, fetched) = 0
                            Application.Current.Dispatcher.Invoke(
                                        Sub()
                                            result.Add(New TreeViewFolder(Me, pidl(0), Me, _setIsLoadingAction))
                                        End Sub)
                        End While
                    End If
                End If

                Me.Folders = result
                Me.NotifyOfPropertyChange("FoldersThreaded")
            End If

            Return _folders
        End Get
        Friend Set(value As ObservableCollection(Of TreeViewFolder))
            SetValue(_folders, value)
        End Set
    End Property

    Protected Overrides Sub shell_Notification(sender As Object, e As NotificationEventArgs)
        ' to be implemented
    End Sub
End Class
