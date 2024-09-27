'Imports System.Runtime.InteropServices
'Imports System.Runtime.InteropServices.ComTypes

'Public Class DataObject
'    Implements IDataObject

'    Private data As New Dictionary(Of FORMATETC, STGMEDIUM)
'    Private adviseList As New List(Of AdviseEntry)

'    Public Function GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer Implements IDataObject.GetData
'        Dim cfFormat As Integer = format.cfFormat
'        Dim f As FORMATETC = data.Keys.FirstOrDefault(Function(k) k.cfFormat = cfFormat)
'        If f.cfFormat = cfFormat Then
'            medium = data(format)
'            Return 0 ' S_OK
'        End If
'        Return &H80040064 ' DV_E_FORMATETC
'    End Function

'    Public Function GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) As Integer Implements IDataObject.GetDataHere
'        ' Implementation for GetDataHere
'        Return 0 ' S_OK
'    End Function

'    Public Function QueryGetData(ByRef format As FORMATETC) As Integer Implements IDataObject.QueryGetData
'        Dim cfFormat As Integer = format.cfFormat
'        Dim f As FORMATETC = data.Keys.FirstOrDefault(Function(k) k.cfFormat = cfFormat)
'        If f.cfFormat = cfFormat Then
'            Return 0 ' S_OK
'        End If
'        Return &H80040064 ' DV_E_FORMATETC
'    End Function

'    Public Function GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) As Integer Implements IDataObject.GetCanonicalFormatEtc
'        formatOut = formatIn
'        Return 0 ' S_OK
'    End Function

'    Public Function SetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) As Integer Implements IDataObject.SetData
'        data(format) = medium
'        Return 0 ' S_OK
'    End Function

'    Public Function EnumFormatEtc(ByVal direction As DATADIR, <Out()> ByRef enumFormatEtc2 As IEnumFORMATETC) As Integer Implements IDataObject.EnumFormatEtc
'        enumFormatEtc2 = New EnumFORMATETCer(data.Keys.ToArray())
'        Return 0 ' S_OK
'    End Function

'    Public Function DAdvise(ByRef format As FORMATETC, ByVal advf As ADVF, ByVal adviseSink As IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject.DAdvise
'        Dim entry As New AdviseEntry With {
'            .Format = format,
'            .Advf = advf,
'            .AdviseSink = adviseSink,
'            .Connection = adviseList.Count + 1
'        }
'        adviseList.Add(entry)
'        connection = entry.Connection
'        Return 0 ' S_OK
'    End Function

'    Public Function DUnadvise(ByVal connection As Integer) As Integer Implements IDataObject.DUnadvise
'        Dim entry = adviseList.FirstOrDefault(Function(a) a.Connection = connection)
'        If entry IsNot Nothing Then
'            adviseList.Remove(entry)
'            Return 0 ' S_OK
'        End If
'        Return &H80040004 ' OLE_E_ADVISENOTSUPPORTED
'    End Function

'    Public Function EnumDAdvise(<Out()> ByRef enumAdvise As IEnumSTATDATA) As Integer Implements IDataObject.EnumDAdvise
'        Return HResult.NotImplemented
'    End Function

'    Private Class AdviseEntry
'        Public Property Format As FORMATETC
'        Public Property Advf As ADVF
'        Public Property AdviseSink As IAdviseSink
'        Public Property Connection As Integer
'    End Class

'    Private Class EnumFORMATETCer
'        Implements IEnumFORMATETC

'        Private formats As FORMATETC()
'        Private currentIndex As Integer

'        Public Sub New(formats As FORMATETC())
'            Me.formats = formats
'            Me.currentIndex = 0
'        End Sub

'        Public Function [Next](celt As Integer, rgelt() As FORMATETC, ByRef pceltFetched As Integer) As Integer Implements IEnumFORMATETC.Next
'            Dim fetched As Integer = 0
'            While currentIndex < formats.Length AndAlso fetched < celt
'                rgelt(fetched) = formats(currentIndex)
'                currentIndex += 1
'                fetched += 1
'            End While
'            pceltFetched = fetched
'            Return If(fetched = celt, 0, 1) ' S_OK or S_FALSE
'        End Function

'        Public Function Skip(celt As Integer) As Integer Implements IEnumFORMATETC.Skip
'            currentIndex += celt
'            Return If(currentIndex <= formats.Length, 0, 1) ' S_OK or S_FALSE
'        End Function

'        Public Function Reset() As Integer Implements IEnumFORMATETC.Reset
'            currentIndex = 0
'            Return 0
'        End Function

'        Public Function Clone(ByRef ppenum As IEnumFORMATETC) As Integer Implements IEnumFORMATETC.Clone
'            ppenum = New EnumFORMATETCer(formats)
'            Return 0 ' S_OK
'        End Function
'    End Class
'End Class