Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes
Imports System.Runtime.Serialization
Imports System.Text
Imports System.Windows

Public Class ShellItem2DataObject
    Implements IDataObject

    Private dsObjectNames As DSOBJECTNAMES

    Private Const DV_E_FORMATETC As Int32 = &H80040064
    Private Const S_OK As Int32 = &H0

    Private pdsObjNames As IntPtr

    Public Sub New(shellItems As List(Of IShellItem2))
        ' Initialize the DSOBJECTNAMES structure
        dsObjectNames = New DSOBJECTNAMES()
        dsObjectNames.clsidNamespace = New Guid("9ac9fbe1-e0a2-4ad6-b4ee-e212013ea917") ' Set appropriate CLSID
        dsObjectNames.cItems = shellItems.Count
        ReDim dsObjectNames.aObjects(shellItems.Count - 1)

        Dim baseOffset As Integer = Marshal.SizeOf(GetType(DSOBJECTNAMES))
        For i As Integer = 0 To shellItems.Count - 1
            dsObjectNames.aObjects(i) = New DSOBJECT()
            dsObjectNames.aObjects(i).dwFlags = 0 ' Example flag
            dsObjectNames.aObjects(i).dwProviderFlags = 0 ' Example provider flag
            dsObjectNames.aObjects(i).offsetName = baseOffset + (i * 4096) ' Example offset for name
            dsObjectNames.aObjects(i).offsetClass = baseOffset + (i * 4096) + 2048 ' Example offset for class
        Next

        ' Allocate memory for the DSOBJECTNAMES structure
        Dim totalSize As Integer = baseOffset + (shellItems.Count * 248)
        pdsObjNames = Marshal.AllocHGlobal(totalSize)

        ' Copy the structure to the allocated memory
        Marshal.StructureToPtr(dsObjectNames, pdsObjNames, False)

        ' Set the pwszName values
        For i As Integer = 0 To shellItems.Count - 1
            Dim name As String
            shellItems(i).GetDisplayName(SHGDN.FORPARSING, name)
            Dim nameBytes As Byte() = Encoding.Unicode.GetBytes(name & vbNullChar)
            Dim pwszName As IntPtr = IntPtr.Add(pdsObjNames, dsObjectNames.aObjects(i).offsetName)
            Marshal.Copy(nameBytes, 0, pwszName, nameBytes.Length)

            Dim class1 As String = "{9ac9fbe1-e0a2-4ad6-b4ee-e212013ea917}"
            Dim classBytes As Byte() = Encoding.Unicode.GetBytes(class1 & vbNullChar)
            Dim pwszclass As IntPtr = IntPtr.Add(pdsObjNames, dsObjectNames.aObjects(i).offsetClass)
            Marshal.Copy(classBytes, 0, pwszclass, classBytes.Length)
        Next
    End Sub

    Public Sub GetData(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetData
        Debug.WriteLine(DataFormats.GetDataFormat(format.cfFormat).ToString())
        If format.cfFormat = DataFormats.GetDataFormat("CFSTR_DSOBJECTNAMES").Id Then
            ' Allocate global memory for DSOBJECTNAMES structure
            'Dim hGlobal As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(dsObjectNames))
            'Marshal.StructureToPtr(dsObjectNames, hGlobal, False)
            medium.tymed = TYMED.TYMED_HGLOBAL
            medium.unionmember = pdsObjNames
            medium.pUnkForRelease = Nothing
        Else
            Throw New COMException("Unsupported format", DV_E_FORMATETC)
        End If
    End Sub

    ' Implement other IDataObject methods as needed
    Public Sub GetDataHere(ByRef format As FORMATETC, ByRef medium As STGMEDIUM) Implements IDataObject.GetDataHere
        Throw New NotImplementedException()
    End Sub

    Public Function QueryGetData(ByRef format As FORMATETC) As Integer Implements IDataObject.QueryGetData
        If format.cfFormat = DataFormats.GetDataFormat("CFSTR_DSOBJECTNAMES").Id Then
            Return S_OK
        Else
            Return DV_E_FORMATETC
        End If
    End Function
    Public Sub GetCanonicalFormatEtc(ByRef formatIn As FORMATETC, ByRef formatOut As FORMATETC) Implements IDataObject.GetCanonicalFormatEtc
        Throw New NotImplementedException()
    End Sub

    Public Sub SetData(ByRef formatIn As FORMATETC, ByRef medium As STGMEDIUM, ByVal release As Boolean) Implements IDataObject.SetData
        Throw New NotImplementedException()
    End Sub

    Public Function EnumFormatEtc(ByVal direction As Integer) As IEnumFORMATETC Implements IDataObject.EnumFormatEtc
        Throw New NotImplementedException()
    End Function

    Public Function DAdvise(ByRef format As FORMATETC, ByVal advf As Integer, ByVal adviseSink As IAdviseSink, ByRef connection As Integer) As Integer Implements IDataObject.DAdvise
        Throw New NotImplementedException()
    End Function

    Public Sub DUnadvise(ByVal connection As Integer) Implements IDataObject.DUnadvise
        Throw New NotImplementedException()
    End Sub

    Public Function EnumDAdvise() As IEnumSTATDATA Implements IDataObject.EnumDAdvise
        Throw New NotImplementedException()
    End Function
End Class
