﻿Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

<ComImport>
<Guid("886D8EEB-8CF2-4446-8D02-CDBA1DBDCF99")>
<InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
Public Interface IPropertyStore
	'Inherits IPersistStream
	Function GetCount(ByRef propertyCount As UInt32) As HResult

	Function GetAt(<[In]> propertyIndex As UInt32, ByRef key As PROPERTYKEY) As HResult

	Function GetValue(<[In]> ByRef key As PROPERTYKEY, <[Out]> ByRef pv As PROPVARIANT) As HResult

	Function SetValue(<[In]> ByRef key As PROPERTYKEY, <[In]> pv As PROPVARIANT) As HResult

	Function Commit() As HResult
End Interface
