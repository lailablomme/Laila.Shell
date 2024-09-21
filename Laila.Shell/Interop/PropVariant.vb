Public Structure PROPVARIANT
	Implements IDisposable

	Public vt As Short

	Public wReserved1 As UInt16

	Public wReserved2 As UInt16

	Public wReserved3 As UInt16

	Public union As PropVariantUnion

	Public Sub Dispose() Implements IDisposable.Dispose
		Functions.PropVariantClear(Me)
	End Sub
End Structure
