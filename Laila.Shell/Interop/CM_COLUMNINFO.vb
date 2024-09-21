Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Public Structure CM_COLUMNINFO
	Public cbSize As UInt32

	Public dwMask As UInt32

	Public dwState As CM_STATE

	Public uWidth As UInt32

	Public uDefaultWidth As UInt32

	Public uIdealWidth As UInt32

	<MarshalAs(UnmanagedType.ByValTStr, SizeConst:=79)>
	Public wszName As String
End Structure
