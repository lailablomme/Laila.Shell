'Imports System.Runtime.InteropServices

'Private Shared Sub addDropDescription(dataObject As ComTypes.IDataObject, dwEffect As DROPEFFECT)
'    'Dim i As Integer = Functions.RegisterClipboardFormat("DropDescription")
'    'While i > Short.MaxValue
'    '    i -= 65536  
'    'End While

'    'Dim formatEtc As New FORMATETC With {
'    '    .cfFormat = i,
'    '    .ptd = IntPtr.Zero,
'    '    .dwAspect = DVASPECT.DVASPECT_CONTENT,
'    '    .lindex = -1,
'    '    .tymed = TYMED.TYMED_HGLOBAL
'    '}

'    'Dim dropDescription As DROPDESCRIPTION
'    'If dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_COPY) Then
'    '    dropDescription.szMessage = "Copy here"
'    '    dropDescription.type = DropImageType.DROPIMAGE_COPY
'    'ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_MOVE) Then
'    '    dropDescription.szMessage = "Move here"
'    '    dropDescription.type = DropImageType.DROPIMAGE_MOVE
'    'ElseIf dwEffect.HasFlag(DROPEFFECT.DROPEFFECT_LINK) Then
'    '    dropDescription.szMessage = "Create shortcut here"
'    '    dropDescription.type = DropImageType.DROPIMAGE_LINK
'    'Else
'    '    dropDescription.szMessage = ""
'    '    dropDescription.type = DropImageType.DROPIMAGE_NONE
'    'End If
'    'dropDescription.szInsert = ""

'    'Dim ptr As IntPtr = Marshal.AllocHGlobal(Marshal.SizeOf(GetType(DROPDESCRIPTION)))
'    'Marshal.StructureToPtr(dropDescription, ptr, False)

'    'Dim medium As System.Runtime.InteropServices.ComTypes.STGMEDIUM
'    'medium.pUnkForRelease = IntPtr.Zero
'    'medium.tymed = ComTypes.TYMED.TYMED_HGLOBAL
'    'medium.unionmember = ptr

'    'dataObject.SetData(formatEtc, medium, True)
'End Sub
