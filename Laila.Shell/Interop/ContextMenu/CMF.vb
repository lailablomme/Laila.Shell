﻿Namespace Interop.ContextMenu
    <Flags>
    Public Enum CMF
        CMF_NORMAL = &H0
        CMF_DEFAULTONLY = &H1
        CMF_VERBSONLY = &H2
        CMF_EXPLORE = &H4
        CMF_NOVERBS = &H8
        CMF_CANRENAME = &H10
        CMF_NODEFAULT = &H20
        CMF_INCLUDESTATIC = &H40
        CMF_ITEMMENU = &H80
        CMF_EXTENDEDVERBS = &H100
        CMF_OPTIMIZEFORINVOKE = &H800
        CMF_SYNCCASCADEMENU = &H1000
        CMF_RESERVED = &HFFFF0000
    End Enum
End Namespace
