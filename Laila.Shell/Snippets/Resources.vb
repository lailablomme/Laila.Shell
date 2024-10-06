'Dim hInstance As IntPtr = Functions.GetModuleHandle("shell32.dll")
'Dim copyString As New System.Text.StringBuilder(256)
'Dim cutString As New System.Text.StringBuilder(256)
'Dim pasteString As New System.Text.StringBuilder(256)
'Dim IDM_COPY As Integer = &HE2 ' Copy command ID
'Dim IDM_CUT As Integer = &H8001  ' Cut command ID
'Dim IDM_PASTE As Integer = &H8002 ' Paste command ID

'' Load localized strings for Copy, Cut, and Paste
'For i = 0 To &HFFFFFF
'Functions.LoadStringW(hInstance, i, copyString, copyString.Capacity)
'If copyString.ToString.Contains("Paste") Then
'Debug.WriteLine(copyString.ToString & "=" & i)
'End If
'Next
