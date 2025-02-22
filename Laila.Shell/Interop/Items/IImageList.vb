Imports System.Runtime.InteropServices
Imports System.Windows

Namespace Interop.Items
    <ComImport>
    <InterfaceType(ComInterfaceType.InterfaceIsIUnknown)>
    <Guid("46eb5926-582e-4017-9fdf-e8998daa0950")>
    Public Interface IImageList
        <PreserveSig>
        Function Add(hbmImage As IntPtr, hbmMask As IntPtr, ByRef pi As Integer) As Integer

        <PreserveSig>
        Function ReplaceIcon(i As Integer, hicon As IntPtr, ByRef pi As Integer) As Integer

        <PreserveSig>
        Function SetOverlayImage(iImage As Integer, iOverlay As Integer) As Integer

        <PreserveSig>
        Function Replace(i As Integer, hbmImage As IntPtr, hbmMask As IntPtr) As Integer

        <PreserveSig>
        Function AddMasked(hbmImage As IntPtr, crMask As Integer, ByRef pi As Integer) As Integer

        <PreserveSig>
        Function Draw(pimldp As IMAGELISTDRAWPARAMS) As Integer

        <PreserveSig>
        Function Remove(i As Integer) As Integer

        <PreserveSig>
        Function GetIcon(i As Integer, flags As Integer, ByRef phicon As IntPtr) As Integer

        <PreserveSig>
        Function GetImageInfo(i As Integer, ByRef pImageInfo As IMAGEINFO) As Integer

        <PreserveSig>
        Function Copy(iDst As Integer, punkSrc As IImageList, iSrc As Integer, uFlags As Integer) As Integer

        <PreserveSig>
        Function Merge(i1 As Integer, punk2 As IImageList, i2 As Integer, dx As Integer, dy As Integer, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer

        <PreserveSig>
        Function Clone(ByRef riid As Guid, ByRef ppv As IntPtr) As Integer
        <PreserveSig>
        Function GetImageRect(i As Integer, ByRef prc As Rect) As Integer

        <PreserveSig>
        Function GetIconSize(ByRef cx As Integer, ByRef cy As Integer) As Integer

        <PreserveSig>
        Function SetIconSize(cx As Integer, cy As Integer) As Integer

        <PreserveSig>
        Function GetImageCount(ByRef pi As Integer) As Integer

        <PreserveSig>
        Function SetImageCount(uNewCount As Integer) As Integer

        <PreserveSig>
        Function SetBkColor(clrBk As Integer, ByRef pclr As Integer) As Integer

        <PreserveSig>
        Function GetBkColor(ByRef pclr As Integer) As Integer

        <PreserveSig>
        Function BeginDrag(iTrack As Integer, dxHotspot As Integer, dyHotspot As Integer) As Integer

        <PreserveSig>
        Function EndDrag() As Integer

        <PreserveSig>
        Function DragEnter(hwndLock As IntPtr, x As Integer, y As Integer) As Integer

        <PreserveSig>
        Function DragLeave(hwndLock As IntPtr) As Integer

        <PreserveSig>
        Function DragMove(x As Integer, y As Integer) As Integer

        <PreserveSig>
        Function SetDragCursorImage(punk As IImageList, iDrag As Integer, dxHotspot As Integer, dyHotspot As Integer) As Integer

        <PreserveSig>
        Function DragShowNolock(fShow As Integer) As Integer
        <PreserveSig>
        Function GetDragImage(ByRef ppt As POINT, ByRef pptHotspot As POINT, ByRef riid As Guid, ByRef ppv As IntPtr) As Integer

        <PreserveSig>
        Function GetItemFlags(i As Integer, ByRef dwFlags As Integer) As Integer

        <PreserveSig>
        Function GetOverlayImage(iOverlay As Integer, ByRef piIndex As Integer) As Integer
    End Interface
End Namespace
