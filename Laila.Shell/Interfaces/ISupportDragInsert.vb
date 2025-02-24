Imports Laila.Shell.Interop
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Interfaces
    Public Interface ISupportDragInsert
        Function DragInsertBefore(dataObject As IDataObject, files As List(Of Item), index As Integer) As HRESULT
        Function Drop(dataObject As IDataObject, files As List(Of Item), index As Integer) As HRESULT
    End Interface
End Namespace