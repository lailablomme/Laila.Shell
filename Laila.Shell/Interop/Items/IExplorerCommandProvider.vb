Imports System.Runtime.InteropServices
Imports Laila.Shell.Interop.Windows

Namespace Interop.Items
    <Guid("64961751-0835-43C0-8FFE-D57686530E64")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <ComImport>
    Public Interface IExplorerCommandProvider
        <PreserveSig>
        Function GetCommands(
        <[In]> pUnkSite As IntPtr,
        <[In]> ByRef riid As Guid,
        <Out> ByRef ppv As IEnumExplorerCommand) As Integer

        <PreserveSig>
        Function GetCommand(
        <[In]> ByRef rGuid As Guid,
        <Out> ByRef ppcmd As IExplorerCommand) As Integer
    End Interface

    <Guid("A88826F8-186F-4987-AADE-EA0CEF8FBFE8")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <ComImport>
    Public Interface IEnumExplorerCommand
        <PreserveSig>
        Function [Next](
        <[In]> celt As UInteger,
        <Out> <MarshalAs(UnmanagedType.Interface)> ByRef pUICommand As IExplorerCommand,
        <Out> ByRef pceltFetched As UInteger) As Integer

        <PreserveSig>
        Function Skip(<[In]> celt As UInteger) As Integer

        <PreserveSig>
        Function Reset() As Integer

        <PreserveSig>
        Function Clone(<Out> ByRef ppenum As IEnumExplorerCommand) As Integer
    End Interface

    <Guid("A08CE4D0-FA25-44AB-B57C-C7B1C323E0B9")>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <ComImport>
    Public Interface IExplorerCommand
        <PreserveSig>
        Function GetTitle(
        <[In], MarshalAs(UnmanagedType.Interface)> ppsiItemArray As IShellItemArray,
        <Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String) As Integer

        <PreserveSig>
        Function GetIcon(
        <[In], MarshalAs(UnmanagedType.Interface)> ppsiItemArray As IShellItemArray,
        <Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszIcon As String) As Integer

        <PreserveSig>
        Function GetToolTip(
        <[In], MarshalAs(UnmanagedType.Interface)> ppsiItemArray As IShellItemArray,
        <Out> <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszInfotip As String) As Integer

        <PreserveSig>
        Function GetCanonicalName(
        <Out> ByRef pguidCommandName As Guid) As Integer

        <PreserveSig>
        Function GetState(
        <[In], MarshalAs(UnmanagedType.Interface)> ppsiItemArray As IShellItemArray,
        <[In]> fOkToBeSlow As Boolean,
        <Out> ByRef pCmdState As EXPCMDSTATE) As Integer

        <PreserveSig>
        Function Invoke(
        <[In], MarshalAs(UnmanagedType.Interface)> ppsiItemArray As IShellItemArray,
        <[In]> pbc As IntPtr) As Integer

        <PreserveSig>
        Function GetFlags(<Out> ByRef pFlags As EXPCMDFLAGS) As Integer

        <PreserveSig>
        Function EnumSubCommands(<Out> ByRef ppenum As IEnumExplorerCommand) As Integer
    End Interface

    ' Enum for command state
    <Flags>
    Public Enum EXPCMDSTATE As UInt32
        ECS_ENABLED = &H0
        ECS_DISABLED = &H1
        ECS_HIDDEN = &H2
        ECS_DISABLEDVISIBLE = &H3
    End Enum
    <Guid("2d4ba2cd-3f9f-4d6c-9feb-5f7f97ad1ee4"),
 InterfaceType(ComInterfaceType.InterfaceIsIUnknown),
 ComImport>
    Public Interface IExplorerCommandArray
        <PreserveSig> Function GetCommands(<Out> ByRef ppEnum As IntPtr) As Integer
    End Interface
    <Flags>
    Public Enum EXPCMDFLAGS As Integer
        ECF_DEFAULT = &H0 ' No command flags set
        ECF_HASSUBCOMMANDS = &H1 ' The command has subcommands
        ECF_HASSPLITBUTTON = &H2 ' A split button is displayed
        ECF_HIDELABEL = &H4 ' The label is hidden
        ECF_ISSEPARATOR = &H8 ' The command is a separator
        ECF_HASLUASHIELD = &H10 ' A UAC shield is displayed
        ECF_SEPARATORBEFORE = &H20 ' The command is located immediately below a separator
        ECF_SEPARATORAFTER = &H40 ' The command is located immediately above a separator
        ECF_ISDROPDOWN = &H80 ' Selecting the command opens a drop-down submenu
        ECF_TOGGLEABLE = &H100 ' Toggleable command (introduced in Windows 8)
        ECF_AUTOMENUICONS = &H200 ' Automatically show menu icons (introduced in Windows 8)
    End Enum
End Namespace