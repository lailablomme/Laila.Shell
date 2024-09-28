Imports System.ComponentModel.Design
Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComVisible(True)>
Public Class DragDataObject
    Implements ComTypes.IDataObject
    Implements IDisposable

    Private _data As Dictionary(Of Short, Tuple(Of FORMATETC, STGMEDIUM, Boolean)) =
        New Dictionary(Of Short, Tuple(Of FORMATETC, STGMEDIUM, Boolean))()
    Private _toRelease As List(Of STGMEDIUM) = New List(Of STGMEDIUM)()
    Private disposedValue As Boolean

    Public Sub GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetData
        If _data.ContainsKey(format.cfFormat) Then
            medium = _data(format.cfFormat).Item2

            Select Case format.cfFormat
                'Case ClipboardFormat.CF_HDROP,
                '     Functions.RegisterClipboardFormat("ComputedImage"),
                '     Functions.RegisterClipboardFormat("DisableDragTextX"),
                '     Functions.RegisterClipboardFormat("DragContextX"),
                '     Functions.RegisterClipboardFormat("DragImageBitsX"),
                '     Functions.RegisterClipboardFormat("DragSourceHelperFlagsX"),
                '     Functions.RegisterClipboardFormat("DragWindow"),
                '     Functions.RegisterClipboardFormat("DropDescription"),
                '     Functions.RegisterClipboardFormat("IsComputingImage"),
                '     Functions.RegisterClipboardFormat("IsShowingLayered"),
                '     Functions.RegisterClipboardFormat("IsShowingText"),
                '     Functions.RegisterClipboardFormat("UsingDefaultDragImageX")
                Case Else
                    If _data(format.cfFormat).Item3 Then
                        _data.Remove(format.cfFormat)
                        '_toRelease.Add(medium)
                    End If
            End Select
        End If
    End Sub

    Public Sub GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetDataHere
    End Sub

    Public Function QueryGetData(ByRef format As FORMATETC) As Integer Implements IDataObject.QueryGetData
        If _data.ContainsKey(format.cfFormat) Then
            Return 0 ' S_OK
        Else
            Return &H80040064 ' DV_E_FORMATETC
        End If
    End Function

    Public Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer Implements IDataObject.GetCanonicalFormatEtc
        formatOut = formatIn
        Return &H80040130 ' DATA_S_SAMEFORMATETC
    End Function

    Public Sub SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) Implements IDataObject.SetData
        If _data.Keys.Contains(format.cfFormat) Then
            Dim m As STGMEDIUM = _data(format.cfFormat).Item2
            If _data(format.cfFormat).Item3 Then
                _toRelease.Add(m)
            End If
            _data(format.cfFormat) = New Tuple(Of FORMATETC, STGMEDIUM, Boolean)(format, medium, release)
        Else
            _data.Add(format.cfFormat, New Tuple(Of FORMATETC, STGMEDIUM, Boolean)(format, medium, release))
        End If
    End Sub

    Public Function EnumFormatEtc(ByVal direction As DATADIR) As IEnumFORMATETC Implements IDataObject.EnumFormatEtc
        Return New EnumFORMATETC(_data.Values.Select(Function(v) v.Item1).ToArray())
    End Function

    Public Function DAdvise(ByRef pFormatetc As FORMATETC, advf As ADVF, adviseSink As IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject.DAdvise
        Return &H80004001 ' E_NOTIMPL
    End Function

    Public Sub DUnadvise(ByVal connection As Integer) Implements IDataObject.DUnadvise
    End Sub

    Public Function EnumDAdvise(ByRef enumAdvise As IEnumSTATDATA) As Integer Implements IDataObject.EnumDAdvise
        Return Nothing
    End Function

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' dispose managed state (managed objects)
            End If

            ' free unmanaged resources (unmanaged objects) and override finalizer
            ' set large fields to null
            Try
                For Each m In _toRelease
                    If m.tymed = TYMED.TYMED_HGLOBAL Then
                        Functions.ReleaseStgMedium(m)
                    End If
                Next
            Catch ex As Exception
            End Try
            disposedValue = True
        End If
    End Sub

    ' override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
    Protected Overrides Sub Finalize()
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=False)
        MyBase.Finalize()
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Public Class EnumFORMATETC
    Implements IEnumFORMATETC

    Private formats As FORMATETC()
    Private currentIndex As Integer

    Public Sub New(formats As FORMATETC())
        Me.formats = formats
        Me.currentIndex = 0
    End Sub

    Public Function [Next](celt As Integer, rgelt() As FORMATETC, pceltFetched() As Integer) As Integer Implements IEnumFORMATETC.Next
        Dim fetched As Integer = 0
        While currentIndex < formats.Length AndAlso fetched < celt
            rgelt(fetched) = formats(currentIndex)
            currentIndex += 1
            fetched += 1
        End While
        ReDim pceltFetched(0)
        pceltFetched(0) = fetched
        Return If(fetched = celt, 0, 1) ' S_OK or S_FALSE
    End Function

    Public Function Skip(celt As Integer) As Integer Implements IEnumFORMATETC.Skip
        currentIndex += celt
        Return If(currentIndex <= formats.Length, 0, 1) ' S_OK or S_FALSE
    End Function

    Public Function Reset() As Integer Implements IEnumFORMATETC.Reset
        currentIndex = 0
    End Function

    Public Sub Clone(ByRef ppenum As IEnumFORMATETC) Implements IEnumFORMATETC.Clone
        ppenum = New EnumFORMATETC(formats)
    End Sub
End Class