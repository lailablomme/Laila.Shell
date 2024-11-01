﻿Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows.Forms

Public Class Clipboard
    Private Shared ReadOnly GMEM_MOVEABLE As Integer = &H2

    Public Shared Sub CopyFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject = New DragDataObject()

        ClipboardFormats.CFSTR_SHELLIDLIST.SetData(dataObject, items)
        ClipboardFormats.CF_HDROP.SetData(dataObject, items)
        If Not items.ToList().Exists(Function(i) TypeOf i Is Folder) Then
            ClipboardFormats.CFSTR_FILEDESCRIPTOR.SetData(dataObject, items)
            ClipboardFormats.CFSTR_FILECONTENTS.SetData(dataObject, items)
        End If
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_COPY)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Sub CutFiles(items As IEnumerable(Of Item))
        Dim dataObject As IDataObject = New DragDataObject()

        ClipboardFormats.CFSTR_SHELLIDLIST.SetData(dataObject, items)
        ClipboardFormats.CF_HDROP.SetData(dataObject, items)
        If Not items.ToList().Exists(Function(i) TypeOf i Is Folder) Then
            ClipboardFormats.CFSTR_FILEDESCRIPTOR.SetData(dataObject, items)
            ClipboardFormats.CFSTR_FILECONTENTS.SetData(dataObject, items)
        End If
        ClipboardFormats.CFSTR_PREFERREDDROPEFFECT.SetData(dataObject, DROPEFFECT.DROPEFFECT_MOVE)

        Functions.OleSetClipboard(dataObject)
    End Sub

    Public Shared Function GetFileNameList(dataObj As IDataObject) As String()
        Dim files() As String
        files = ClipboardFormats.CFSTR_SHELLIDLIST.GetData(dataObj)?.Select(Function(i) i.FullPath).ToArray()
        If files Is Nothing OrElse files.Count = 0 Then
            files = ClipboardFormats.CF_HDROP.GetData(dataObj)
        End If
        Return files
    End Function

    Public Shared Function GetHasGlobalData(clipboardFormat As String)
        Dim dataObject As IDataObject
        Functions.OleGetClipboard(dataObject)
        Return GetHasGlobalData(dataObject, clipboardFormat)
    End Function

    Public Shared Function GetHasGlobalData(dataObject As IDataObject, clipboardFormat As String)
        Return GetHasGlobalData(dataObject, Functions.RegisterClipboardFormat(clipboardFormat))
    End Function

    Public Shared Function GetHasGlobalData(dataObject As IDataObject, clipboardFormat As Short)
        Dim format As FORMATETC = New FORMATETC() With {
            .cfFormat = clipboardFormat,
            .dwAspect = DVASPECT.DVASPECT_CONTENT,
            .lindex = -1,
            .ptd = IntPtr.Zero,
            .tymed = TYMED.TYMED_HGLOBAL
        }
        Return dataObject.QueryGetData(format) = 0
    End Function
End Class
