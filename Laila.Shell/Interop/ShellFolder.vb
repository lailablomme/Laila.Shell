Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Public Class ShellFolder
    Implements IShellFolder

    Private _sf As IShellFolder
    Private _dataObj As IDataObject

    Public Sub New(sf As IShellFolder, dataObj As IDataObject)
        _sf = sf
        _dataObj = dataObj
    End Sub

    Public Function ParseDisplayName(hwndOwner As Integer, pbcReserved As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> lpszDisplayName As String, ByRef pchEaten As Integer, ByRef ppidl As IntPtr, ByRef pdwAttributes As Integer) As Integer Implements IShellFolder.ParseDisplayName
        Return _sf.ParseDisplayName(hwndOwner, pbcReserved, lpszDisplayName, pchEaten, ppidl, pdwAttributes)
    End Function

    Public Function EnumObjects(hwndOwner As Integer, <MarshalAs(UnmanagedType.U4)> grfFlags As SHCONTF, ByRef ppenumIDList As IEnumIDList) As Integer Implements IShellFolder.EnumObjects
        Return _sf.EnumObjects(hwndOwner, grfFlags, ppenumIDList)
    End Function

    Public Function BindToObject(pidl As IntPtr, pbcReserved As IntPtr, ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer Implements IShellFolder.BindToObject
        Return _sf.BindToObject(pidl, pbcReserved, riid, ppvOut)
    End Function

    Public Function BindToStorage(pidl As IntPtr, pbcReserved As IntPtr, ByRef riid As Guid, ppvObj As IntPtr) As Integer Implements IShellFolder.BindToStorage
        Return _sf.BindToStorage(pidl, pbcReserved, riid, ppvObj)
    End Function

    Public Function CompareIDs(lParam As UInteger, pidl1 As IntPtr, pidl2 As IntPtr) As Integer Implements IShellFolder.CompareIDs
        Return _sf.CompareIDs(lParam, pidl1, pidl2)
    End Function

    Public Function CreateViewObject(hwndOwner As IntPtr, ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer Implements IShellFolder.CreateViewObject
        Return _sf.CreateViewObject(hwndOwner, riid, ppvOut)
    End Function

    Public Function GetAttributesOf(cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> apidl() As IntPtr, ByRef rgfInOut As Integer) As Integer Implements IShellFolder.GetAttributesOf
        Return _sf.GetAttributesOf(cidl, apidl, rgfInOut)
    End Function

    Public Function GetUIObjectOf(hwndOwner As IntPtr, cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> apidl() As IntPtr, ByRef riid As Guid, ByRef prgfInOut As Integer, ByRef ppvOut As IntPtr) As Integer Implements IShellFolder.GetUIObjectOf
        If riid.Equals(GetType(IDataObject).GUID) Then
            'Dim dataObj As OurDataObject = New OurDataObject(New System.Windows.DataObject())
            'Dim ienum As IEnumFORMATETC = _dataObj.EnumFormatEtc(DATADIR.DATADIR_GET)
            'Dim fetched(0) As Integer
            'fetched(0) = 1
            'Do While fetched(0) = 1
            '    Dim f(0) As FORMATETC
            '    ienum.Next(1, f, fetched)
            '    For Each ff In f
            '        Debug.WriteLine(System.Windows.DataFormats.GetDataFormat(ff.cfFormat).Name)
            '    Next
            'Loop
            'Dim fm As FORMATETC
            ''fm.dwAspect = DVASPECT.DVASPECT_CONTENT
            ''fm.cfFormat = System.Windows.DataFormats.GetDataFormat("Format0").Id
            'fm.tymed = TYMED.TYMED_FILE
            'Dim med As STGMEDIUM
            '_dataObj.GetData(fm, med)
            'dataObj.SetData(med)
            Dim ptr As IntPtr
            Try
                ptr = Marshal.GetIUnknownForObject(_dataObj)
                Return Marshal.QueryInterface(ptr, riid, ppvOut)
            Finally
                Marshal.Release(ptr)
            End Try
        Else
            Return _sf.GetUIObjectOf(hwndOwner, cidl, apidl, riid, prgfInOut, ppvOut)
        End If
    End Function

    Public Function GetDisplayNameOf(pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> uFlags As SHGDN, lpName As IntPtr) As Integer Implements IShellFolder.GetDisplayNameOf
        Return _sf.GetDisplayNameOf(pidl, uFlags, lpName)
    End Function

    Public Function SetNameOf(hwndOwner As Integer, pidl As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> lpszName As String, <MarshalAs(UnmanagedType.U4)> uFlags As SHGDN, ByRef ppidlOut As IntPtr) As Integer Implements IShellFolder.SetNameOf
        Return _sf.SetNameOf(hwndOwner, pidl, lpszName, uFlags, ppidlOut)
    End Function
End Class
