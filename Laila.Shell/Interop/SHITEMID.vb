Public Structure SHITEMID
    ''' <summary>
    ''' The size of identifier, in bytes, including cb itself.
    ''' </summary>
    Public cb As UShort
    ''' <summary>
    ''' A variable-length item identifier.
    ''' </summary>
    Public abID() As Byte
End Structure
