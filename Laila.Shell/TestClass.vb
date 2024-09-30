Imports System.Runtime.InteropServices
Imports System.Windows

<ComVisible(True), Guid("cab7cbe3-aec5-468b-89ae-ffb73ba6223f"), ProgId("Laila.Shell.TestClass")>
Public Class TestClass
    <ComVisible(True)>
    Public Sub Test()
        MessageBox.Show("Test")
    End Sub
End Class
