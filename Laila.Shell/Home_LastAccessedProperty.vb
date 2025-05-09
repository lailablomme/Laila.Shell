Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Public Class Home_LastAccessedProperty
    Inherits [Property]

    Public Overloads Shared Property Key As PROPERTYKEY = New PROPERTYKEY("390DCE80-0981-4BBC-A55D-95F189A5A8FD:2")
    Public Shared Shadows Property CanonicalName As String = "Laila.Shell.Home.LastAccessed"

    Public Sub New()
        MyBase.New(New CachedPropertyDescription() With {
                      .PropertyKey = Home_LastAccessedProperty.Key,
                      .CanonicalName = Home_LastAccessedProperty.CanonicalName
                   }, CType(Nothing, IShellItem2))
        _isCustom = True
        _displayName = My.Resources.Folder_Home_LastAccessedPropertyDisplayName
        _dataType = VarEnum.VT_FILETIME
    End Sub

    Public Sub New(value As DateTime?)
        Me.New()
        _rawValue = New PROPVARIANT()
        If value.HasValue Then _rawValue.SetValue(value.Value)
    End Sub
End Class
