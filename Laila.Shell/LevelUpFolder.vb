Imports System.Runtime.InteropServices
Imports System.Windows.Media
Imports System.Windows.Media.Imaging

Public Class LevelUpFolder
    Inherits Folder

    Private _parentFolder As Folder

    Public Overloads Shared Function FromParsingName(parsingName As String, parent As Folder, setIsLoadingAction As Action(Of Boolean)) As Folder
        Dim ptr As IntPtr, ptr2 As IntPtr
        Functions.SHCreateItemFromParsingName(parsingName, IntPtr.Zero, Guids.IID_IShellItem2, ptr)
        Dim shellItem2 As IShellItem2 = Marshal.GetTypedObjectForIUnknown(ptr, GetType(IShellItem2))
        shellItem2.BindToHandler(Nothing, Guids.BHID_SFObject, GetType(IShellFolder).GUID, ptr2)
        Dim shellFolder As IShellFolder = Marshal.GetTypedObjectForIUnknown(ptr2, GetType(IShellFolder))
        Return New LevelUpFolder(shellFolder, shellItem2, parent, setIsLoadingAction)
    End Function

    Public Sub New(shellFolder As IShellFolder, shellItem2 As IShellItem2, parentFolder As Folder, setIsLoadingAction As Action(Of Boolean))
        MyBase.New(shellFolder, shellItem2, setIsLoadingAction)
        _parentFolder = parentFolder
    End Sub

    Public Overrides ReadOnly Property Icon16 As ImageSource
        Get
            Return New ImageSourceConverter().ConvertFromString("pack://application:,,,/Laila.Shell;component/Images/arrow_up16.png")
        End Get
    End Property

    Public Overrides ReadOnly Property Overlay16 As ImageSource
        Get
            Return Nothing
        End Get
    End Property

    'Public ReadOnly Property PropertyValues(canonicalName As String) As Object
    '    Get
    '        If canonicalName = "System.ItemNameDisplay" Then
    '            Return ".."
    '        Else
    '            Dim propVariant As PROPVARIANT
    '            Me.ShellItem2.GetProperty(_parentFolder.Columns(canonicalName).PROPERTYKEY, propVariant)
    '            Return getPropertyValue(propVariant, canonicalName)
    '        End If
    '    End Get
    'End Property
End Class
