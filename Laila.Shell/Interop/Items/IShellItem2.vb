Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

<ComImport>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
<Guid("7E9FB0D3-919F-4307-AB2E-9B1860310C93")>
Public Interface IShellItem2
    <PreserveSig>
    Function BindToHandler(<[In]> pbc As IntPtr, <[In]> ByRef rbhid As Guid, <[In]> ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer

    '<PreserveSig>
    'Function BindToHandler(<[In], MarshalAs(UnmanagedType.Interface)> ByRef pbc As ComTypes.IBindCtx, <[In]> ByRef rbhid As Guid, <[In]> ByRef riid As Guid, ByRef ppvOut As IntPtr) As Integer

    <PreserveSig>
    Function GetParent(<MarshalAs(UnmanagedType.Interface)> ByRef ppsi As IShellItem) As Integer

    <PreserveSig>
    Function GetDisplayName(<[In]> sigdnName As UInteger, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String) As Integer

    <PreserveSig>
    Sub GetAttributes(<[In]> sfgaoMask As Integer, ByRef psfgaoAttribs As Integer)

    <PreserveSig>
    Sub Compare(<[In], MarshalAs(UnmanagedType.Interface)> psi As IShellItem, <[In]> hint As UInt32, ByRef piOrder As Integer)

    <PreserveSig>
    Function GetPropertyStore(<[In]> flags As Integer, <[In]> ByRef riid As Guid, ByRef ppv As IntPtr) As Integer

    <PreserveSig>
    Sub GetPropertyStoreWithCreateObject(<[In]> flags As Integer, <[In], MarshalAs(UnmanagedType.IUnknown)> punkCreateObject As Object, <[In]> ByRef riid As Guid, ByRef ppv As IntPtr)

    <PreserveSig>
    Sub GetPropertyStoreForKeys(<[In]> ByRef rgKeys As PROPERTYKEY, <[In]> cKeys As UInt32, <[In]> flags As Integer, <[In]> ByRef riid As Guid, <MarshalAs(UnmanagedType.IUnknown)> ByRef ppv As Object)

    <PreserveSig>
    Sub GetPropertyDescriptionList(<[In]> ByRef keyType As PROPERTYKEY, <[In]> ByRef riid As Guid, <Out> ByRef ppv As IntPtr)

    <PreserveSig>
    Function Update(<[In]> pbc As IntPtr) As Integer

    <PreserveSig>
    Sub GetProperty(<[In]> ByRef key As PROPERTYKEY, ByRef ppropvar As PROPVARIANT)

    <PreserveSig>
    Sub GetCLSID(<[In]> ByRef key As PROPERTYKEY, ByRef pclsid As Guid)

    <PreserveSig>
    Sub GetFileTime(<[In]> ByRef key As PROPERTYKEY, ByRef pft As FILETIME)

    <PreserveSig>
    Sub GetInt32(<[In]> ByRef key As PROPERTYKEY, ByRef pi As Integer)

    <PreserveSig>
    Function GetString(<[In]> ByRef key As PROPERTYKEY, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppsz As String) As Integer

    <PreserveSig>
    Sub GetUInt64(<[In]> ByRef key As PROPERTYKEY, ByRef pull As ULong)

    <PreserveSig>
    Sub GetBool(<[In]> ByRef key As PROPERTYKEY, ByRef pf As Integer)
End Interface
'NAME:BHID_AssociationArray VALUE : bea9ef17-82.0F1-4f60-9284-4f8db75c3be9
'NAME:BHID_CollectionFactory VALUE : 9b9eb7fa-cbf0-4db3-ac58-d3a522991b2a
'NAME:BHID_DataObject VALUE : b8c0bd9f-ed24 - 455c-83e6-d5390c4fe8c4
'NAME:BHID_EnumAssocHandlers VALUE : b8ab0b9c-c2ec - 4.0F7a-918d-314900e6280a
'NAME:BHID_EnumAssocHandlersAll VALUE : a4f44537-5d1c-466b-a74a-62937c38fa0c
'NAME:BHID_EnumItems VALUE : 94f60519-2850-4924-aa5a-d15e84868039
'NAME:BHID_FilePlaceholder VALUE : 8677dceb-aae0-4005-8d3d-547fa852f825
'NAME:BHID_Filter VALUE : 38d08778-f557-4690-9ebf-ba54706ad8f7
'NAME:BHID_Item VALUE : a9cf13da-b3ff - 4Dfe-8f66-a383a2b9c911
'NAME:BHID_LinkTargetItem VALUE : 3981e228-f559-11d3-8e3a-00c04f6837d5
'NAME:BHID_LocalizableItem VALUE : 06acee39-54cc-465c-ac07-cefa3f32c9ea
'NAME:BHID_PropertyStore VALUE : 0384e1a4-1523-439c-a4c8-ab911052f586
'NAME:BHID_RandomAccessStream VALUE : f16fc93b-77ae-4cfe-bda7-a866eea6878d
'NAME:BHID_SFObject VALUE : 3981e224-f559-11d3-8e3a-00c04f6837d5
'NAME:BHID_SFUIObject VALUE : 3981e225-f559-11d3-8e3a-00c04f6837d5
'NAME:BHID_SFViewObject VALUE : 3981e226-f559-11d3-8e3a-00c04f6837d5
'NAME:BHID_ShellFolderCommands VALUE : a86ef9df-2a78-4080-8791-9d55d137a2cc
'NAME:BHID_Storage VALUE : 3981e227-f559-11d3-8e3a-00c04f6837d5
'NAME:BHID_StorageEnum VALUE : 4621a4e3-f0d6-4773-8a9c-46e77b174840
'NAME:BHID_StorageItem VALUE : 404e2109-77d2-4699-a5a0-4fdf10db9837
'NAME:BHID_Stream VALUE : 1cebb3ab-7c10-499a-a417-92ca16c4cb83
'NAME:BHID_ThumbnailHandler VALUE : 7b2e650a-8e20-4f4a-b09e-6597afc72fb0
'NAME:BHID_Transfer VALUE : d5e346a1-f753 - 4932 - b403 - 4574800E2498