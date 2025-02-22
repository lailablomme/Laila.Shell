
Imports System
Imports System.Runtime.InteropServices
Imports System.Text

Namespace Interop.Items
    <ComVisible(True), ComImport(),
    InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
    Guid("000214F9-0000-0000-C000-000000000046")>
    Public Interface IShellLinkW
        ''' <summary>Retrieves the path and file name of a Shell link object</summary>
        Sub GetPath(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszFile As StringBuilder,
                cch As Integer,
                <Out(), [Optional]> ByRef pfd As WIN32_FIND_DATAW,
                fFlags As UInteger)

        ''' <summary>Retrieves the list of item identifiers for a Shell link object</summary>
        Sub GetIDList(<Out()> ByRef ppidl As IntPtr)

        ''' <summary>Sets the pointer to an item identifier list (PIDL) for a Shell link object.</summary>
        Sub SetIDList(pidl As IntPtr)

        ''' <summary>Retrieves the description string for a Shell link object</summary>
        Sub GetDescription(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszName As StringBuilder,
                       cch As Integer)

        ''' <summary>Sets the description for a Shell link object</summary>
        Sub SetDescription(<MarshalAs(UnmanagedType.LPWStr)> pszName As String)

        ''' <summary>Retrieves the name of the working directory for a Shell link object</summary>
        Sub GetWorkingDirectory(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszDir As StringBuilder,
                            cch As Integer)

        ''' <summary>Sets the name of the working directory for a Shell link object</summary>
        Sub SetWorkingDirectory(<MarshalAs(UnmanagedType.LPWStr)> pszDir As String)

        ''' <summary>Retrieves the command-line arguments associated with a Shell link object</summary>
        Sub GetArguments(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszArgs As StringBuilder,
                     cch As Integer)

        ''' <summary>Sets the command-line arguments for a Shell link object</summary>
        Sub SetArguments(<MarshalAs(UnmanagedType.LPWStr)> pszArgs As String)

        ''' <summary>Retrieves the hotkey for a Shell link object</summary>
        Sub GetHotkey(<Out()> ByRef pwHotkey As Short)

        ''' <summary>Sets a hotkey for a Shell link object</summary>
        Sub SetHotkey(wHotkey As Short)

        ''' <summary>Retrieves the show command for a Shell link object</summary>
        Sub GetShowCmd(<Out()> ByRef piShowCmd As Integer)

        ''' <summary>Sets the show command for a Shell link object</summary>
        Sub SetShowCmd(iShowCmd As Integer)

        ''' <summary>Retrieves the location (path and index) of the icon for a Shell link object</summary>
        Sub GetIconLocation(<Out(), MarshalAs(UnmanagedType.LPWStr)> pszIconPath As StringBuilder,
                        cch As Integer,
                        <Out()> ByRef piIcon As Integer)

        ''' <summary>Sets the location (path and index) of the icon for a Shell link object</summary>
        Sub SetIconLocation(<MarshalAs(UnmanagedType.LPWStr)> pszIconPath As String,
                        iIcon As Integer)

        ''' <summary>Sets the relative path to the Shell link object</summary>
        Sub SetRelativePath(<MarshalAs(UnmanagedType.LPWStr)> pszPathRel As String,
                        dwReserved As UInteger)

        ''' <summary>Attempts to find the target of a Shell link, even if it has been moved or renamed</summary>
        Sub Resolve(hwnd As IntPtr, fFlags As SLR_FLAGS)

        ''' <summary>Sets the path and file name of a Shell link object</summary>
        Sub SetPath(<MarshalAs(UnmanagedType.LPWStr)> pszFile As String)
    End Interface


    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
    Public Structure WIN32_FIND_DATAW
        Public dwFileAttributes As UInteger
        Public ftCreationTime As Long
        Public ftLastAccessTime As Long
        Public ftLastWriteTime As Long
        Public nFileSizeHigh As UInteger
        Public nFileSizeLow As UInteger
        Public dwReserved0 As UInteger
        Public dwReserved1 As UInteger
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=260)>
        Public cFileName As String
        <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=14)>
        Public cAlternateFileName As String
    End Structure

    ' SLGP_FLAGS: Flags for GetPath
    <Flags()>
    Public Enum SLGP_FLAGS
        ShortPath = &H1
        UNCPRIORITY = &H2
        RAWPATH = &H4
        RELATIVEPRIORITY = &H8
    End Enum

    ' SLR_FLAGS: Flags for Resolve
    <Flags()>
    Public Enum SLR_FLAGS
        NO_UI = &H1
        ANY_MATCH = &H2
        UPDATE = &H4
        NOUPDATE = &H8
        NOSEARCH = &H10
        NOTRACK = &H20
        NORELINK = &H40
        INVOKE_MSI = &H80
    End Enum
End Namespace

