Imports Laila.Shell.Interop
Imports Laila.Shell.Interop.DragDrop
Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Controls

Namespace Interfaces
    Public Interface ISupportDragInsert
        Function DragInsertBefore(dataObject As IDataObject_PreserveSig, files As List(Of Item), index As Integer, overListBoxItem As ListBoxItem) As HRESULT
        Function Drop(dataObject As IDataObject_PreserveSig, files As List(Of Item), index As Integer) As HRESULT
        ReadOnly Property Items As ObservableCollection(Of Item)
    End Interface
End Namespace