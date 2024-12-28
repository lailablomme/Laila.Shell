Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows

<ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("dfd3b6b5-c10c-4be9-85f6-a66969f402f6")>
Public Interface IExplorerBrowser
    ' Prepares the browser to be navigated
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub Initialize(
        <[In]> hwndParent As IntPtr,
        <[In]> ByRef prc As WIN32RECT,
        <[Optional]> ByRef pfs As FOLDERSETTINGS
    )

    ' Destroys the browser
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub Destroy()

    ' Sets the size and position of the view windows created by the browser
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub SetRect(
        <[In], Out> phdwp As IntPtr,
        <[In]> ByRef rcBrowser As Rect
    )

    ' Sets the name of the property bag
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub SetPropertyBag(
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszPropertyBag As String
    )

    ' Sets the default empty text
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub SetEmptyText(
        <[In], MarshalAs(UnmanagedType.LPWStr)> pszEmptyText As String
    )

    ' Sets the folder settings for the current view
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub SetFolderSettings(<[In]> ByRef pfs As FOLDERSETTINGS)

    ' Initiates a connection with IExplorerBrowser for event callbacks
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub Advise(
        <[In]> psbe As IntPtr,
        <Out> ByRef pdwCookie As UInteger
    )

    ' Terminates an advisory connection
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub Unadvise(<[In]> dwCookie As UInteger)

    ' Sets the current browser options
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub SetOptions(<[In]> dwFlag As EXPLORER_BROWSER_OPTIONS)

    ' Gets the current browser options
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetOptions(<Out> ByRef pdwFlag As EXPLORER_BROWSER_OPTIONS)

    ' Browses to a pointer to an item identifier list (PIDL)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub BrowseToIDList(<[In]> pidl As IntPtr, <[In]> uFlags As SBSP)

    ' Browses to an object
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime), PreserveSig>
    Function BrowseToObject(
        <MarshalAs(UnmanagedType.IUnknown)> punk As Object,
        uFlags As SBSP
    ) As HRESULT

    ' Creates a results folder and fills it with items
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub FillFromObject(
        <[In], MarshalAs(UnmanagedType.IUnknown)> punk As Object,
        <[In]> dwFlags As EXPLORER_BROWSER_FILL_FLAGS
    )

    ' Removes all items from the results folder
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub RemoveAll()

    ' Gets an interface for the current view of the browser
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetCurrentView(<[In]> ByRef riid As Guid) As <MarshalAs(UnmanagedType.Interface, IidParameterIndex:=0)> Object
End Interface

''' <summary>These flags are used with IExplorerBrowser::FillFromObject.</summary>
<Flags>
Public Enum EXPLORER_BROWSER_FILL_FLAGS
    ''' <summary>No flags.</summary>
    EBF_NONE = &H0

    ''' <summary>
    ''' Causes IExplorerBrowser::FillFromObject to first populate the results folder with the contents of the parent folders of the
    ''' items in the data object, and then select only the items that are in the data object.
    ''' </summary>
    EBF_SELECTFROMDATAOBJECT = &H100

    ''' <summary>
    ''' Do not allow dropping on the folder. In other words, do not register a drop target for the view. Applications can then
    ''' register their own drop targets.
    ''' </summary>
    EBF_NODROPTARGET = &H200
End Enum

''' <summary>These flags are used with IExplorerBrowser::GetOptions and IExplorerBrowser::SetOptions.</summary>
<Flags>
Public Enum EXPLORER_BROWSER_OPTIONS
    ''' <summary>No options.</summary>
    EBO_NONE = &H0

    ''' <summary>Do not navigate further than the initial navigation.</summary>
    EBO_NAVIGATEONCE = &H1

    ''' <summary>
    ''' Use the following standard panes: Commands Module pane, Navigation pane, Details pane, and Preview pane. An implementer of
    ''' IExplorerPaneVisibility can modify the components of the Commands Module that are shown. For more information see,
    ''' IExplorerPaneVisibility::GetPaneState. If EBO_SHOWFRAMES is not set, Explorer browser uses a single view object.
    ''' </summary>
    EBO_SHOWFRAMES = &H2

    ''' <summary>Always navigate, even if you are attempting to navigate to the current folder.</summary>
    EBO_ALWAYSNAVIGATE = &H4

    ''' <summary>Do not update the travel log.</summary>
    EBO_NOTRAVELLOG = &H8

    ''' <summary>
    ''' Do not use a wrapper window. This flag is used with legacy clients that need the browser parented directly on themselves.
    ''' </summary>
    EBO_NOWRAPPERWINDOW = &H10

    ''' <summary>Show WebView for SharePoint sites.</summary>
    EBO_HTMLSHAREPOINTVIEW = &H20

    ''' <summary>Introduced in Windows Vista. Do not draw a border around the browser window.</summary>
    EBO_NOBORDER = &H40

    ''' <summary>Introduced in Windows Vista. Do not persist the view state.</summary>
    EBO_NOPERSISTVIEWSTATE = &H80
End Enum

''' <summary>
''' Flags for Shell Browser Navigation.
''' </summary>
<Flags>
Public Enum SBSP As Integer
    ''' <summary>Default browser.</summary>
    DEFBROWSER = &H0

    ''' <summary>Same browser instance.</summary>
    SAMEBROWSER = &H1

    ''' <summary>New browser instance.</summary>
    NEWBROWSER = &H2

    ''' <summary>Default navigation mode.</summary>
    DEFMODE = &H0

    ''' <summary>Open mode.</summary>
    OPENMODE = &H10

    ''' <summary>Explore mode.</summary>
    EXPLOREMODE = &H20

    ''' <summary>Help mode.</summary>
    HELPMODE = &H40

    ''' <summary>Do not transfer history.</summary>
    NOTRANSFERHIST = &H80

    ''' <summary>Absolute navigation.</summary>
    ABSOLUTE = &H0

    ''' <summary>Relative navigation.</summary>
    RELATIVE = &H1000

    ''' <summary>Navigate to parent folder.</summary>
    PARENT = &H2000

    ''' <summary>Navigate back.</summary>
    NAVIGATEBACK = &H4000

    ''' <summary>Navigate forward.</summary>
    NAVIGATEFORWARD = &H8000

    ''' <summary>Allow auto-navigation.</summary>
    ALLOW_AUTONAVIGATE = &H10000

    ''' <summary>Keep the same template. Available in Windows Vista and later.</summary>
    KEEPSAMETEMPLATE = &H20000

    ''' <summary>Keep word wheel text. Available in Windows Vista and later.</summary>
    KEEPWORDWHEELTEXT = &H40000

    ''' <summary>Activate without focus. Available in Windows Vista and later.</summary>
    ACTIVATE_NOFOCUS = &H80000

    ''' <summary>Create no history. Available in Windows Vista and later.</summary>
    CREATENOHISTORY = &H100000

    ''' <summary>Play no sound. Available in Windows Vista and later.</summary>
    PLAYNOSOUND = &H200000

    ''' <summary>Caller is untrusted. Available in IE6 SP2 and later.</summary>
    CALLERUNTRUSTED = &H800000

    ''' <summary>Trust first download. Available in IE6 SP2 and later.</summary>
    TRUSTFIRSTDOWNLOAD = &H1000000

    ''' <summary>Untrusted for download. Available in IE6 SP2 and later.</summary>
    UNTRUSTEDFORDOWNLOAD = &H2000000

    ''' <summary>No auto-select.</summary>
    NOAUTOSELECT = &H4000000

    ''' <summary>Write no history.</summary>
    WRITENOHISTORY = &H8000000

    ''' <summary>Trusted for ActiveX. Available in IE6 SP2 and later.</summary>
    TRUSTEDFORACTIVEX = &H10000000

    ''' <summary>Feed navigation. Available in IE7 and later.</summary>
    FEEDNAVIGATION = &H20000000

    ''' <summary>Redirect.</summary>
    REDIRECT = &H40000000

    ''' <summary>Initiated by hyperlink frame.</summary>
    INITIATEDBYHLINKFRAME = &H80000000
End Enum