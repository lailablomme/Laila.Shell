Imports System.Runtime.InteropServices

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
        _displayName = "Home category"
        _dataType = VarEnum.VT_UI4
        _displayType = PropertyDisplayType.String
    End Sub

    Public Sub New(value As [Type])
        Me.New()
        _rawValue = New PROPVARIANT()
        _rawValue.SetValue(Convert.ToUInt32(value))
    End Sub

    Public Enum [Type]
        PINNED_ITEM = 3
        FREQUENT_FOLDER = 2
        RECENT_FILE = 1
    End Enum

    Public Overrides ReadOnly Property Text As String
        Get
            Select Case CType(Me.Value, [Type])
                Case [Type].PINNED_ITEM
                    Return "Pinned items"
                Case [Type].FREQUENT_FOLDER
                    Return "Frequent folders"
                Case [Type].RECENT_FILE
                    Return "Recent files"
                Case Else
                    Return "Unknown"
            End Select
        End Get
    End Property
End Class
