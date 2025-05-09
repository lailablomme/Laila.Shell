Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Items
Imports Laila.Shell.Interop.Properties

Public Class Home_CategoryProperty
    Inherits [Property]

    Public Overloads Shared Property Key As PROPERTYKEY = New PROPERTYKEY("390DCE80-0981-4BBC-A55D-95F189A5A8FD:1")
    Public Shared Shadows Property CanonicalName As String = "Laila.Shell.Home.Category"

    Public Sub New()
        MyBase.New(New CachedPropertyDescription() With {
                      .PropertyKey = Home_CategoryProperty.Key,
                      .CanonicalName = Home_CategoryProperty.CanonicalName
                   }, CType(Nothing, IShellItem2))
        _isCustom = True
        _displayName = My.Resources.Folder_Home_CategoryPropertyDisplayName
        _dataType = VarEnum.VT_UI4
        _displayType = PropertyDisplayType.String
    End Sub

    Public Sub New(value As [Type])
        Me.New()
        _rawValue = New PROPVARIANT()
        _rawValue.SetValue(Convert.ToUInt32(value))
    End Sub

    Public Enum [Type]
        QUICK_LAUNCH = 3
        FAVORITES = 2
        RECENT_FILE = 1
    End Enum

    Public Overrides ReadOnly Property Text As String
        Get
            Select Case CType(Me.Value, [Type])
                Case [Type].QUICK_LAUNCH
                    Return My.Resources.Folder_Home_QuickLaunch
                Case [Type].FAVORITES
                    Return "Favorites"
                Case [Type].RECENT_FILE
                    Return My.Resources.Folder_Home_RecentFiles
                Case Else
                    Return "Unknown"
            End Select
        End Get
    End Property
End Class
