Imports System.Collections.Specialized
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows

Public Class DragDrop
    Public Shared Sub DoDragDrop(source As DependencyObject, items As IEnumerable(Of Item))
        Dim dataObject As DataObject = New DataObject()
        Dim files As StringCollection = New StringCollection()
        For Each i In items
            files.Add(i.FullPath)
        Next
        dataObject.SetFileDropList(files)

        System.Windows.DragDrop.DoDragDrop(source, dataObject, DragDropEffects.All)
    End Sub
End Class
