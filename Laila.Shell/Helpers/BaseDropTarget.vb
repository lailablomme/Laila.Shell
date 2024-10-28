Imports System.Runtime.InteropServices.ComTypes

Public MustInherit Class BaseDropTarget
    Implements IDropTarget

    Public MustOverride Function DragEnter(pDataObj As IDataObject, grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragEnter

    Public MustOverride Function DragOver(grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.DragOver

    Public MustOverride Function DragLeave() As Integer Implements IDropTarget.DragLeave

    Public MustOverride Function Drop(pDataObj As IDataObject, grfKeyState As MK, pt As WIN32POINT, ByRef pdwEffect As Integer) As Integer Implements IDropTarget.Drop
End Class
