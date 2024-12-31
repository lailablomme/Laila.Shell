Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Windows
Imports System.Windows.Shapes
Imports System.Runtime.InteropServices.ComTypes
Imports System.IO
Imports Shell32
Imports Laila.Shell.Controls

Public Class Functions

    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function RegSetValueEx(
        hKey As IntPtr,
        lpValueName As String,
        Reserved As Integer,
        dwType As Integer,
        lpData As Byte(),
        cbData As Integer) As Integer
    End Function
    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function RegOpenKeyEx(
        hKey As Integer,
        lpSubKey As String,
        ulOptions As Integer,
        samDesired As Integer,
        ByRef phkResult As IntPtr) As Integer
    End Function
    <DllImport("advapi32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function RegQueryValueEx(
        hKey As IntPtr,
        lpValueName As String,
        lpReserved As Integer,
        ByRef lpType As Integer,
        ByRef lpData As Int32,
        ByRef lpcbData As Integer) As Integer
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Public Shared Function RegCloseKey(hKey As IntPtr) As Integer
    End Function
    <DllImport("advapi32.dll", SetLastError:=True)>
    Public Shared Function RegNotifyChangeKeyValue(
        hKey As IntPtr,
        watchSubtree As Boolean,
        notifyFilter As Integer,
        hEvent As IntPtr,
        async As Boolean) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function SHCreateShellFolderView(
        ByVal pcsfv As SFV_CREATE,
        <Out> ByRef ppsv As IShellView
    ) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function MoveWindow(hWnd As IntPtr, X As Integer, Y As Integer, nWidth As Integer, nHeight As Integer, bRepaint As Boolean) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetParent(hWndChild As IntPtr, hWndNewParent As IntPtr) As IntPtr
    End Function
    <DllImport("shell32.dll", SetLastError:=True)>
    Public Shared Sub SHGetSetSettings(
    ByRef value As SHELLSTATE,
    dwMask As SSF,
    bSet As Boolean
)
    End Sub
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function SendMessageTimeout(
    hWnd As IntPtr,
    Msg As UInteger,
    wParam As IntPtr,
    lParam As IntPtr,
    flags As UInteger,
    timeout As UInteger,
    ByRef result As IntPtr
) As IntPtr
    End Function

    Public Shared Property HWND_BROADCAST As IntPtr = CType(&HFFFF, IntPtr)
    Public Const WM_SETTINGCHANGE As UInteger = &H1A
    Public Const SMTO_ABORTIFHUNG As UInteger = &H2
    Public Const STGM_READWRITE As Integer = 2
    Public Const STR_ENUM_ITEMS_FLAGS As String = "EnumItemsFlags"
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ToUnicode(
        wVirtKey As UInteger,
        wScanCode As UInteger,
        lpKeyState As Byte(),
        <MarshalAs(UnmanagedType.LPWStr)> pwszBuff As String,
        cchBuff As Integer,
        wFlags As UInteger) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetKeyboardState(lpKeyState As Byte()) As Boolean
    End Function
    <DllImport("user32.dll")>
    Public Shared Function MapVirtualKey(uCode As UInteger, uMapType As UInteger) As UInteger
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function SHGetNameFromIDList(ByVal pidl As IntPtr, ByVal sigdnName As SIGDN, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String) As Integer
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Sub GetLocalTime(ByRef lpSystemTime As SYSTEMTIME)
    End Sub
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ShowWindow(hwnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function
    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function AssocQueryStringW(
        flags As AssocF,
        str As AssocStr,
        pszAssoc As String,
        pszExtra As String,
        pszOut As String,
        ByRef pcchOut As UInteger
    ) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetClipboardData(uFormat As UInteger) As IntPtr
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function ILClone(pidl As IntPtr) As IntPtr
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function ILRemoveLastID(pidl As IntPtr) As Boolean
    End Function
    <DllImport("shell32.dll", EntryPoint:="#727", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetImageListHandle(iImageList As Integer, ByRef riid As Guid, ByRef phImageList As IntPtr) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHGetImageList(
        iImageList As Integer,
        ByRef riid As Guid,
        <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IImageList
    ) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetFileInfo(
        pszPath As String,
        dwFileAttributes As Integer,
        ByRef psfi As SHFILEINFO,
        cbFileInfo As Integer,
        uFlags As Integer
    ) As IntPtr
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function OleGetClipboard(ByRef dataObject As System.Runtime.InteropServices.ComTypes.IDataObject) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function ShellExecute(hWnd As IntPtr, lpOperation As String, lpFile As String, lpParameters As String, lpDirectory As String, nShowCmd As Integer) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
    End Function
    <DllImport("user32.dll")>
    Public Shared Function GetShellWindow() As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SetVolumeLabelW(
    <MarshalAs(UnmanagedType.LPWStr)> lpRootPathName As String,
    <MarshalAs(UnmanagedType.LPWStr)> lpVolumeName As String
) As Boolean
    End Function
    <DllImport("ole32.dll", SetLastError:=True)>
    Public Shared Function OleSetClipboard(ByVal pDataObj As IDataObject) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ILCreateFromPath(ByVal pszPath As String) As IntPtr
    End Function
    <DllImport("ole32.dll", CharSet:=CharSet.Unicode, SetLastError:=True, ExactSpelling:=True)>
    Public Shared Function StgCreateStorageEx(
    ByVal pwcsName As String,
    ByVal grfMode As STGM,
    ByVal stgfmt As STGFMT,
    ByVal grfAttrs As UInteger,
    ByVal pSecurityDescriptor As IntPtr,
    ByVal reserved As UInteger,
    ByVal riid As Guid,
    <Out()> ByRef ppObject As IntPtr
) As Integer
    End Function
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function GetFileAttributesW(
        <MarshalAs(UnmanagedType.LPWStr)> ByVal lpFileName As String
    ) As UInteger
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto, SetLastError:=True, BestFitMapping:=False, ThrowOnUnmappableChar:=True)>
    Public Shared Sub SHAddToRecentDocs(ByVal uFlags As UInt32, <MarshalAs(UnmanagedType.LPWStr)> ByVal path As String)
    End Sub
    <DllImport("propsys.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function PSGetNameFromPropertyKey(ByRef propkey As PROPERTYKEY, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszCanonicalName As String) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHCreateItemArray(
        ByVal cItems As UInteger,
        <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=0)> ByVal rgpsi As IShellItem(),
        ByVal punkAttributes As IntPtr,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppsia As IShellItemArray) As Integer
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Sub RtlMoveMemory(ByVal dest As IntPtr, ByVal src As IntPtr, ByVal count As UIntPtr)
    End Sub
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function GlobalSize(ByVal hMem As IntPtr) As UIntPtr
    End Function
    <DllImport("shell32.dll", EntryPoint:="DragQueryFileW", CharSet:=CharSet.Auto)>
    Public Shared Function DragQueryFile(hDrop As IntPtr, iFile As UInteger, lpFile As StringBuilder, cch As UInteger) As UInteger
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function DragFinish(hDrop As IntPtr) As Boolean
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHFileOperation(ByRef FileOp As SHFILEOPSTRUCT) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function RegisterDragDrop(hwnd As IntPtr, pDropTarget As IDropTarget) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function RevokeDragDrop(ByVal hwnd As IntPtr) As Integer
    End Function
    '<DllImport("propsys.dll", SetLastError:=True, CharSet:=CharSet.Unicode)>
    'Public Shared Function PSCreateMemoryPropertyStore(
    '    ByRef riid As Guid,
    '    ByRef ppv As IntPtr
    ') As Integer
    'End Function
    <DllImport("kernel32.dll")>
    Public Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function
    <DllImport("ole32.dll")>
    Public Shared Sub ReleaseStgMedium(ByRef pmedium As STGMEDIUM)
    End Sub
    <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function LoadStringW(hInstance As IntPtr, uID As UInteger, lpBuffer As StringBuilder, nBufferMax As Integer) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetClipboardData(uFormat As Integer, hMem As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyCursor(hCursor As IntPtr) As Boolean
    End Function
    Public Const CFSTR_PREFERREDDROPEFFECT As String = "Preferred DropEffect"
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function EmptyClipboard() As Boolean
    End Function
    <DllImport("user32.dll", EntryPoint:="RegisterClipboardFormat", SetLastError:=True, CharSet:=CharSet.Auto)>
    Public Shared Function RegisterClipboardFormatWIN32(lpString As String) As UInteger
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function LoadCursor(hInstance As IntPtr, lpCursorName As String) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetWindowLong(hwnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetWindowLong(hwnd As IntPtr, nIndex As Integer) As Integer
    End Function
    Public Shared Function RegisterClipboardFormat(lpString As String) As Short
        Dim i As Integer = Functions.RegisterClipboardFormatWIN32(lpString)
        While i > Short.MaxValue
            i -= 65536
        End While
        Return i
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Public Shared Function GetFileAttributesEx(
        ByVal lpFileName As String,
        ByVal fInfoLevelId As Integer,
        ByRef lpFileInformation As WIN32_FILE_ATTRIBUTE_DATA) As Boolean
    End Function
    <DllImport("Kernel32.dll", SetLastError:=True)>
    Public Shared Function SystemTimeToFileTime(ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function OpenClipboard(hWndNewOwner As IntPtr) As Boolean
    End Function

    <DllImport("shlwapi.dll", CallingConvention:=CallingConvention.StdCall, PreserveSig:=True, CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreateStreamOnFileEx(
        <[In]> ByVal pszFile As String,
        <[In]> ByVal grfMode As Integer,
        <[In]> ByVal dwAttributes As UInteger,
        <[In]> ByVal fCreate As UInteger,
        <[In]> ByVal pstmTemplate As IntPtr,
        <Out> ByRef ppstm As IntPtr
    ) As HRESULT
    End Function
    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function SHCreateStreamOnFileW(
        <[In]> ByVal fileName As String,
        <[In]> ByVal grfMode As Integer,
        <Out> ByRef ppstm As IStream
    ) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function CloseClipboard() As Boolean
    End Function

    <DllImport("ole32.dll")>
    Public Shared Function OleCreateDataObject(
    ByRef riid As Guid,
    ByVal pUnkOuter As IntPtr,
    ByVal riidDataObject As Guid,
    ByRef ppvObj As IntPtr
) As Integer
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function GlobalAlloc(uFlags As Integer, dwBytes As Integer) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function GlobalLock(hMem As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function GlobalUnlock(hMem As IntPtr) As Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function GlobalFree(hMem As IntPtr) As Integer
    End Function
    <DllImport("kernel32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function LoadLibrary(lpFileName As String) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)>
    Public Shared Function FreeLibrary(hModule As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetCursorPos(ByRef lpPoint As Point) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetWindowPos _
    (ByVal hwnd As IntPtr, ByVal hWndInsertAfter As IntPtr,
     ByVal x As Integer, ByVal y As Integer,
     ByVal cx As Integer, ByVal cy As Integer,
     ByVal uFlags As UInteger) As Boolean
    End Function
    '<DllImport("user32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    'Public Shared Function RegisterClipboardFormat(lpString As String) As UInteger
    'End Function
    <DllImport("ole32.dll")>
    Public Shared Function DoDragDrop(pDataObj As IDataObject, pDropSource As IDropSource, dwOKEffects As Integer, <Out> ByRef pdwEffect As Integer) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function OleInitialize(ByVal pvReserved As IntPtr) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ILGetSize(ByVal pidl As IntPtr) As UInteger
    End Function
    <DllImport("urlmon.dll")>
    Public Shared Function CopyStgMedium(ByRef pcstgmedSrc As STGMEDIUM, ByRef pstgmedDest As STGMEDIUM) As Integer
    End Function
    '<DllImport("ole32.dll")>
    'Public Shared Function CoCreateInstance(
    '    ByRef clsid As Guid,
    '    ByVal pUnkOuter As IntPtr,
    '    ByVal dwClsContext As UInteger,
    '    ByRef riid As Guid,
    '    <MarshalAs(UnmanagedType.Interface)> ByRef ppv As FolderView.IExplorerHost) As Integer
    'End Function
    <DllImport("ole32.dll")>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IShellFolder) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPreviewHandler) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IFileOperation) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IDropTargetHelper) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CoCreateInstance(
        ByRef clsid As Guid,
        ByVal pUnkOuter As IntPtr,
        ByVal dwClsContext As UInteger,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IDragSourceHelper) As Integer
    End Function
    <DllImport("kernel32.dll", EntryPoint:="RtlMoveMemory", SetLastError:=False)>
    Public Shared Sub CopyMemory(destination As IntPtr, source As IntPtr, length As UInteger)
    End Sub
    <DllImport("ole32.dll")>
    Public Shared Sub OleUninitialize()
    End Sub
    <DllImport("Shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateShellItemArrayFromIDLists(
        ByVal cidl As UInteger,
        ByVal rgpidl As IntPtr(),
        <MarshalAs(UnmanagedType.Interface)> ByRef ppsiItemArray As IShellItemArray) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetDC(hWnd As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateDataObject(pidlFolder As IntPtr, cidl As UInteger, apidl As IntPtr, pdtInner As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IDataObject) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateDataObject(pidlFolder As IntPtr, cidl As UInteger, apidl() As IntPtr, pdtInner As IntPtr, ByRef riid As Guid, <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IDataObject) As Integer
    End Function
    <DllImport("ole32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function CoCreateBindCtx(dwFlags As UInt32, reserved As UInt32, pMalloc As IntPtr, ByRef ppbc As IBindCtx) As Integer
    End Function

    <DllImport("gdi32.dll", SetLastError:=True)>
    Public Shared Function DeleteObject(ByVal hObject As IntPtr) As Boolean
    End Function
    <DllImport("ole32.dll", SetLastError:=True)>
    Public Shared Function PropVariantClear(ByRef pvar As PROPVARIANT) As Integer
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHGetDesktopFolder(ByRef ppshf As IShellFolder) As Integer
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHGetKnownFolderItem(<[In]> rfid As Guid, <[In]> flags As UInteger, <[In]> hToken As IntPtr, <[In]> riid As Guid, <Out> ByRef ppv As IShellItem) As HRESULT
    End Function

    <DllImport("shlwapi.dll")>
    Public Shared Function StrRetToBuf(ByVal pstr As IntPtr, ByVal pidl As IntPtr, ByVal pszBuf As StringBuilder, <MarshalAs(UnmanagedType.U4)> ByVal cchBuf As Integer) As Integer
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function PathParseIconLocationA(<[In], Out, MarshalAs(UnmanagedType.LPStr)> ByRef pszIconFile As String) As Integer
    End Function

    <DllImport("Shell32.dll", EntryPoint:="ExtractIconExW", CharSet:=CharSet.Unicode, ExactSpelling:=True, CallingConvention:=CallingConvention.StdCall)>
    Public Shared Function ExtractIconEx(sFile As String, iIndex As Integer, ByRef piLargeVersion As IntPtr, ByRef piSmallVersion As IntPtr, amountIcons As Integer) As Integer
    End Function

    <DllImport("user32.dll")>
    Public Shared Function GetIconInfo(ByVal hIcon As IntPtr, ByRef piconinfo As ICONINFO) As Boolean
    End Function

    <DllImport("shell32.dll")>
    Public Shared Function SHCreateItemWithParent(pidlParent As IntPtr, psfParent As IShellFolder, pidl As IntPtr, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer
    End Function
    <DllImport("ole32.dll")>
    Public Shared Function CreateBindCtx(reserved As UInt32, ByRef ppbc As IntPtr) As Integer
    End Function

    <DllImport("shlwapi.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreatePropertyBagOnMemory(
    grfMode As Integer,
    ByRef riid As Guid,
    <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyBag) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, PreserveSig:=False)>
    Public Shared Function SHCreateShellItemArrayFromShellItem(
        <MarshalAs(UnmanagedType.Interface)> psi As IShellItem,
        ByRef riid As Guid,
        <MarshalAs(UnmanagedType.Interface)> ByRef ppv As IShellItemArray
    ) As Integer

    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreateItemFromParsingName(pszPath As String, pbc As IntPtr, riid As Guid, ByRef ppv As IntPtr) As HRESULT
    End Function

    <DllImport("shell32.dll")>
    Public Shared Sub SHParseDisplayName(<MarshalAs(UnmanagedType.LPWStr)> name As String, bindingContext As IntPtr, <Out()> ByRef pidl As IntPtr, sfgaoIn As UInt32, <Out()> ByRef sfgaoOut As UInt32)
    End Sub

    <DllImport("shell32.dll")>
    Public Shared Function SHCreateItemFromIDList(<[In]> pidl As IntPtr, <[In]> riid As Guid, <Out> ByRef ppv As IntPtr) As HRESULT
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSGetPropertyDescription(propKey As PROPERTYKEY, riid As Guid, ByRef ppv As IntPtr) As Long
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSGetImageReferenceForValue(<[In]> ByRef propKey As PROPERTYKEY, <[In]> ByRef propvar As PROPVARIANT, <Out> ByRef ppszImageRes As IntPtr) As Long
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function PSCreateMemoryPropertyStore(<[In]> riid As Guid, <Out> ByRef ppv As IntPtr) As HRESULT
    End Function

    <DllImport("propsys.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function PSFormatForDisplay(
    ByRef propkey As PROPERTYKEY,
    ByRef propvar As PROPVARIANT,
    ByVal pdfFlags As PropertyDescriptionFormatOptions,
    ByVal pwszText As StringBuilder,
    ByVal cchText As UInteger) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Unicode)>
    Public Shared Function SHCreateDefaultContextMenu(<[In]> ByRef pdcm As DEFCONTEXTMENU, <[In]> ByRef riid As Guid, <Out> ByRef ppv As IntPtr) As Integer
    End Function
    <DllImport("propsys.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
    Public Shared Function PSGetPropertyDescription(
            ByRef propkey As PROPERTYKEY,
            ByRef riid As Guid,
            <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyDescription
        ) As HRESULT
    End Function

    <DllImport("propsys.dll")>
    Public Shared Function SHLoadPropertyStoreFromBlob(pBlob As PROPVARIANT, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppPropStore As IPropertyStore) As HRESULT
    End Function

    <DllImport("shlwapi.dll", SetLastError:=True)>
    Public Shared Function SHCreateMemStream(pvBuf As IntPtr, ByVal cbSize As Long) As IntPtr
    End Function

    <DllImport("shell32.dll", SetLastError:=True)>
    Public Shared Function SHGetPropertyStoreForFile(
        ByVal pszFilePath As String,
        ByVal grfFlags As UInt32,
        ByRef ppPropStore As IntPtr
    ) As HRESULT
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertyStoreForFile(
            <MarshalAs(UnmanagedType.LPWStr)> pszFilePath As String,
             grfFlags As Integer,
            <MarshalAs(UnmanagedType.Interface)> ByRef ppPropStore As IPropertyStore
        ) As HRESULT
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertySystem(
  <[In]> riid As Guid,
  <[Out]> ByRef ppv As IntPtr
) As HRESULT
    End Function

    <DllImport("propsys.dll", SetLastError:=True)>
    Public Shared Function PSGetPropertyDescriptionByName(
    <[In], MarshalAs(UnmanagedType.LPWStr)> pszCanonicalName As String,
  <[In]> riid As Guid,
  <[Out], MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyDescription
) As HRESULT

    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMessage(ByRef lpMsg As MSG, ByVal hWnd As IntPtr, ByVal wMsgFilterMin As UInteger, ByVal wMsgFilterMax As UInteger) As Boolean
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function SHGetIDListFromObject(ByVal punk As IntPtr, ByRef ppidl As IntPtr) As Integer
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Unicode, ExactSpelling:=True)>
    Public Shared Function SHGetIDListFromObject(<MarshalAs(UnmanagedType.Interface)> punk As IShellItem2, ByRef ppidl As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", SetLastError:=True)>
    Public Shared Function ILCombine(
        ByVal pidl1 As IntPtr,
        ByVal pidl2 As IntPtr
    ) As IntPtr
    End Function
    <DllImport("shell32.dll", SetLastError:=True)>
    Public Shared Function ILFindChild(
        ByVal pidlParent As IntPtr,
        ByVal pidlChild As IntPtr
    ) As IntPtr
    End Function
    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ILIsEqual(ByVal pidl1 As IntPtr, ByVal pidl2 As IntPtr) As Boolean
    End Function
    <DllImport("shell32.dll")>
    Public Shared Function ILFindLastID(pidl As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function CreatePopupMenu() As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuItemCount(hMenu As IntPtr) As Integer
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function Shell_GetImageLists(ByRef phimlLarge As IntPtr, ByRef phimlSmall As IntPtr) As Integer
    End Function

    <DllImport("comctl32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function ImageList_GetIcon(himl As IntPtr, i As Integer, flags As Integer) As IntPtr
    End Function

    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyIcon(ByVal hIcon As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuString(hMenu As IntPtr, uIDItem As UInteger, lpString As StringBuilder, nMaxCount As Integer, uFlag As UInteger) As Integer
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function GetMenuItemInfo(hMenu As IntPtr, uItem As UInteger, fByPosition As Boolean, ByRef lpmii As MENUITEMINFO) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function DestroyMenu(ByVal hMenu As IntPtr) As Boolean
    End Function

    Declare Function TrackPopupMenuEx Lib "user32.dll" (ByVal hMenu As IntPtr,
           ByVal uFlags As Integer, ByVal x As Integer, ByVal y As Integer,
           ByVal hWnd As IntPtr, ByVal lptpm As IntPtr) As Integer
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Sub mouse_event(ByVal dwFlags As Integer, ByVal dx As Integer, ByVal dy As Integer, ByVal dwData As Integer, ByVal dwExtraInfo As IntPtr)
    End Sub
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function EndMenu() As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll", SetLastError:=True)>
    Public Shared Function SetMenu(ByVal hWnd As IntPtr, ByVal hMenu As IntPtr) As Boolean
    End Function

    <DllImport("shell32.dll", CharSet:=CharSet.Auto)>
    Public Shared Function SHGetPathFromIDList(ByVal pidl As IntPtr, ByVal pszPath As System.Text.StringBuilder) As Boolean
    End Function

    <DllImport("shell32", EntryPoint:="#2", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotifyRegister(hwnd As IntPtr, fSources As SHCNRF, fEvents As SHCNE, wMsg As WM, cEntries As Integer, <MarshalAs(UnmanagedType.LPArray)> pfsne As SHChangeNotifyEntry()) As Integer
    End Function

    <DllImport("shell32", EntryPoint:="#4", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotifyDeregister(hNotify As Integer) As Boolean
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotification_Lock(ByVal hChange As IntPtr, ByVal dwProcId As UInt32, ByRef pppidl As IntPtr, ByRef plEvent As SHCNE) As IntPtr
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Function SHChangeNotification_Unlock(ByVal hLock As IntPtr) As Int32
    End Function

    <DllImport("shell32", CharSet:=CharSet.Auto)>
    Public Shared Sub SHChangeNotify(ByVal wEventId As Integer, ByVal uFlags As Integer, ByVal dwItem1 As IntPtr, ByVal dwItem2 As IntPtr)
    End Sub
End Class
