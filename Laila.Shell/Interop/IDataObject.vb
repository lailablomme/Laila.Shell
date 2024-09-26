Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization

<ComVisible(True)>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
<Guid("0000010e-0000-0000-C000-000000000046")>
Public Interface IDataObject2
    Sub GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM)
    Sub GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM)
    Function QueryGetData(ByRef format As FORMATETC) As Integer
    Sub GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC)
    Sub SetData(ByRef formatIn As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean)
    Function EnumFormatEtc(ByVal direction As Integer) As IEnumFORMATETC
    Function DAdvise(ByRef format As FORMATETC, ByVal advf As Integer, ByVal adviseSink As IAdviseSink, ByRef connection As Integer) As Integer
    Sub DUnadvise(ByVal connection As Integer)
    Function EnumDAdvise() As IEnumSTATDATA
End Interface