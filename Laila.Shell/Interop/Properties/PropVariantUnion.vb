Imports System.Runtime.InteropServices
Imports System.Runtime.InteropServices.ComTypes

Namespace Interop.Properties
	<StructLayout(LayoutKind.Explicit, Pack:=1)>
	Public Structure PropVariantUnion
		<FieldOffset(0)>
		Public cVal As SByte

		<FieldOffset(0)>
		Public bVal As Byte
		<FieldOffset(0)>
		Public boolVal As Boolean

		<FieldOffset(0)>
		Public iVal As Short
		<FieldOffset(0)>
		Public uiVal As UShort
		<FieldOffset(0)>
		Public intVal As Integer
		<FieldOffset(0)>
		Public uintVal As UInt32
		<FieldOffset(0)>
		Public hVal As Long
		<FieldOffset(0)>
		Public uhVal As ULong
		<FieldOffset(0)>
		Public fltVal As Single
		<FieldOffset(0)>
		Public dblVal As Double
		<FieldOffset(0)>
		Public cyVal As CY
		<FieldOffset(0)>
		Public filetime As FILETIME
		<FieldOffset(0)>
		Public ptr As IntPtr
		<FieldOffset(0)>
		Public bstrblobVal As BSTRBLOB
		<FieldOffset(0)>
		Public cArray As CArray
	End Structure
End Namespace
