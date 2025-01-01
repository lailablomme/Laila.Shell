Imports System.Runtime.InteropServices

<ComImportAttribute(),
 InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
 Guid("000214E6-0000-0000-C000-000000000046")>
Public Interface IShellFolder
    <PreserveSig()>
    Function ParseDisplayName(ByVal hwndOwner As Integer, ByVal pbcReserved As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszDisplayName As String, ByRef pchEaten As Integer, ByRef ppidl As IntPtr, ByRef pdwAttributes As Integer) As Integer

    <PreserveSig()>
    Function EnumObjects(ByVal hwndOwner As Integer, <MarshalAs(UnmanagedType.U4)> ByVal grfFlags As SHCONTF, ByRef ppenumIDList As IEnumIDList) As Integer

    <PreserveSig()>
    Function BindToObject(ByVal pidl As IntPtr, ByVal pbcReserved As IntPtr, ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer

    <PreserveSig()>
    Function BindToStorage(ByVal pidl As IntPtr, ByVal pbcReserved As IntPtr, ByRef riid As Guid, ByRef ppvObj As IntPtr) As Integer

    <PreserveSig()>
    Function CompareIDs(ByVal lParam As UInt32, ByVal pidl1 As IntPtr, ByVal pidl2 As IntPtr) As Integer

    <PreserveSig()>
    Function CreateViewObject(ByVal hwndOwner As IntPtr, ByRef riid As Guid, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppvOut As IShellView) As Integer

    <PreserveSig()>
    Function GetAttributesOf(ByVal cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal apidl() As IntPtr, ByRef rgfInOut As Integer) As Integer

    <PreserveSig()>
    Function GetUIObjectOf(ByVal hwndOwner As IntPtr, ByVal cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal apidl() As IntPtr, ByRef riid As Guid, ByRef prgfInOut As Integer, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppvOut As IDropTarget) As Integer

    <PreserveSig()>
    Function GetDisplayNameOf(ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> ByVal uFlags As SHGDN, ByVal lpName As IntPtr) As Integer

    <PreserveSig()>
    Function SetNameOf(ByVal hwndOwner As Integer, ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszName As String, <MarshalAs(UnmanagedType.U4)> ByVal uFlags As SHGDN, ByRef ppidlOut As IntPtr) As Integer
End Interface

<ComImportAttribute(),
 InterfaceTypeAttribute(ComInterfaceType.InterfaceIsIUnknown),
 Guid("000214E6-0000-0000-C000-000000000046")>
Public Interface IShellFolderForIContextMenu
    <PreserveSig()>
    Function ParseDisplayName(ByVal hwndOwner As Integer, ByVal pbcReserved As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszDisplayName As String, ByRef pchEaten As Integer, ByRef ppidl As IntPtr, ByRef pdwAttributes As Integer) As Integer

    <PreserveSig()>
    Function EnumObjects(ByVal hwndOwner As Integer, <MarshalAs(UnmanagedType.U4)> ByVal grfFlags As SHCONTF, ByRef ppenumIDList As IEnumIDList) As Integer

    <PreserveSig()>
    Function BindToObject(ByVal pidl As IntPtr, ByVal pbcReserved As IntPtr, ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer

    <PreserveSig()>
    Function BindToStorage(ByVal pidl As IntPtr, ByVal pbcReserved As IntPtr, ByRef riid As Guid, ByRef ppvObj As IntPtr) As Integer

    <PreserveSig()>
    Function CompareIDs(ByVal lParam As UInt32, ByVal pidl1 As IntPtr, ByVal pidl2 As IntPtr) As Integer

    <PreserveSig()>
    Function CreateViewObject(ByVal hwndOwner As IntPtr, ByRef riid As Guid, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppvOut As IContextMenu) As Integer

    <PreserveSig()>
    Function GetAttributesOf(ByVal cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal apidl() As IntPtr, ByRef rgfInOut As Integer) As Integer

    <PreserveSig()>
    Function GetUIObjectOf(ByVal hwndOwner As IntPtr, ByVal cidl As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal apidl() As IntPtr, ByRef riid As Guid, ByRef prgfInOut As Integer, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppvOut As IContextMenu) As Integer

    <PreserveSig()>
    Function GetDisplayNameOf(ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.U4)> ByVal uFlags As SHGDN, ByVal lpName As IntPtr) As Integer

    <PreserveSig()>
    Function SetNameOf(ByVal hwndOwner As Integer, ByVal pidl As IntPtr, <MarshalAs(UnmanagedType.LPWStr)> ByVal lpszName As String, <MarshalAs(UnmanagedType.U4)> ByVal uFlags As SHGDN, ByRef ppidlOut As IntPtr) As Integer
End Interface
