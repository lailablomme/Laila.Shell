Namespace Interop.Windows
    <Flags>
    Public Enum MOUSEEVENTF As UInteger
        MOVE = &H1
        LEFTDOWN = &H2
        LEFTUP = &H4
        RIGHTDOWN = &H8
        RIGHTUP = &H10
        MIDDLEDOWN = &H20
        MIDDLEUP = &H40
        XDOWN = &H80
        XUP = &H100
        WHEEL = &H800
        HWHEEL = &H1000
    End Enum
End Namespace
