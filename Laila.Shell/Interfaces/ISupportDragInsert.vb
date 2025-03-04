Imports Laila.Shell.Interop
Imports System.Collections.ObjectModel
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Windows.Controls

Namespace Interfaces
    Public Interface ISupportDragInsert
        Function DragInsertBefore(dataObject As IDataObject, files As List(Of Item), index As Integer, overListBoxItem As ListBoxItem) As HRESULT
        Function Drop(dataObject As IDataObject, files As List(Of Item), index As Integer) As HRESULT
        ReadOnly Property Items As ObservableCollection(Of Item)
    End Interface
End Namespace