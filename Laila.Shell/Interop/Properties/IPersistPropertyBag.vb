Imports System.Runtime.InteropServices

Namespace Interop.Properties
    <ComImport(), Guid("37D84F60-42CB-11CE-8135-00AA004BB851"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    Public Interface IPersistPropertyBag
        ''' <summary>
        ''' Initializes a new instance of the object.
        ''' </summary>
        <PreserveSig>
        Function InitNew() As Integer

        ''' <summary>
        ''' Loads properties from a property bag.
        ''' </summary>
        ''' <param name="pPropBag">The property bag from which to load properties.</param>
        ''' <param name="pErrorLog">Optional error log, used to log errors during property loading.</param>
        <PreserveSig>
        Function Load(
        <[In], MarshalAs(UnmanagedType.Interface)> ByVal pPropBag As IPropertyBag,
        <[In], MarshalAs(UnmanagedType.Interface)> ByVal pErrorLog As IErrorLog) As Integer

        ''' <summary>
        ''' Saves properties to a property bag.
        ''' </summary>
        ''' <param name="pPropBag">The property bag in which to save properties.</param>
        ''' <param name="fClearDirty">Indicates whether to clear the dirty flag after saving properties.</param>
        ''' <param name="fSaveAllProperties">Indicates whether all properties should be saved, or just the dirty ones.</param>
        <PreserveSig>
        Function Save(
        <[In], MarshalAs(UnmanagedType.Interface)> ByVal pPropBag As IPropertyBag,
        <[In]> ByVal fClearDirty As Boolean,
        <[In]> ByVal fSaveAllProperties As Boolean) As Integer
    End Interface
End Namespace
