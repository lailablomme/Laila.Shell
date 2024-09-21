Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport, Guid("57D2EDED-5062-400E-B107-5DAE79FE57A6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPropertyDescription2
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetPropertyKey(ByRef pkey As PROPERTYKEY)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetCanonicalName(<MarshalAs(UnmanagedType.LPWStr)> ByRef ppszName As String)
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetPropertyType(ByRef pvartype As VarEnum) As HResult
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    <PreserveSig>
    Function GetDisplayName(ByRef ppszName As IntPtr) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetEditInvitation(ByRef ppszInvite As IntPtr) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetTypeFlags(<[In]> mask As PropertyTypeOptions, ByRef ppdtFlags As PropertyTypeOptions) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetViewFlags(ByRef ppdvFlags As PropertyViewOptions) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetDefaultColumnWidth(ByRef pcxChars As UInt32) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetDisplayType(ByRef pdisplaytype As PropertyDisplayType) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetColumnState(ByRef pcsFlags As PropertyColumnStateOptions) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetGroupingRange(ByRef pgr As PropertyGroupingRange) As HResult
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetRelativeDescriptionType(ByRef prdt As RelativeDescriptionType)
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub GetRelativeDescription(<[In]> propvar1 As PROPVARIANT, <[In]> propvar2 As PROPVARIANT, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszDesc1 As String, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszDesc2 As String)
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetSortDescription(ByRef psd As PropertySortDescription) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetSortDescriptionLabel(<[In]> fDescending As Boolean, ByRef ppszDescription As IntPtr) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetAggregationType(ByRef paggtype As PropertyAggregationType) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetConditionType(ByRef pcontype As PropertyConditionType, ByRef popDefault As PropertyConditionOperation) As HResult
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function GetEnumTypeList(<[In]> ByRef riid As Guid, <Out, MarshalAs(UnmanagedType.Interface)> ByRef ppv As IPropertyEnumTypeList) As HResult
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Sub CoerceToCanonicalValue(<[In], Out> propvar As PROPVARIANT)
    <PreserveSig>
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function FormatForDisplay(<[In]> propvar As PROPVARIANT, <[In]> ByRef pdfFlags As PropertyDescriptionFormatOptions, <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszDisplay As String) As HResult
    <MethodImpl(MethodImplOptions.InternalCall, MethodCodeType:=MethodCodeType.Runtime)>
    Function IsValueCanonical(<[In]> propvar As PROPVARIANT) As HResult
    Function GetImageReferenceForValue(<[In]> propvar As PROPVARIANT, <[Out]> <MarshalAs(UnmanagedType.LPWStr)> ByRef ppszImageRes As String) As HResult
End Interface
