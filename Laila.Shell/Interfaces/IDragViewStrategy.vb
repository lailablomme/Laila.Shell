Imports System.Windows
Imports System.Windows.Controls
Imports Laila.Shell.Interop.Windows

Namespace Interfaces
    Public Interface IDragViewStrategy
        Function GetOverListBoxItem(ptWIN32 As WIN32POINT) As ListBoxItem
        Function GetAreAdjecentItems(item1 As ListBoxItem, item2 As ListBoxItem) As Boolean
        Sub GetInsertIndex(ptWIN32 As WIN32POINT, overListBoxItem As ListBoxItem, overItem As Item, ByRef dragInsertParent As ISupportDragInsert, ByRef insertIndex As Integer)
        Sub SetDragInsertIndicator(overListBoxItem As ListBoxItem, overItem As Item, visibility As Visibility, beforeIndex As Integer)
    End Interface
End Namespace