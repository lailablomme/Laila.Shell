'Application.Current.Dispatcher.Invoke(
'    Async Function() As Task
'        Await Task.Delay(1000)

'        For Each item In contextMenu.Items
'            Dim hwndSource As HwndSource = CType(PresentationSource.FromVisual(item), HwndSource)
'            Dim hwnd As IntPtr = hwndSource.Handle
'            Dim measureItem As MEASUREITEMSTRUCT
'            measureItem.CtlType = ODT.ODT_MENU
'            measureItem.CtlID = 0
'            measureItem.itemID = item.Tag.ToString().Split(vbTab)(0)
'            measureItem.itemWidth = CType(item, Control).ActualWidth
'            measureItem.itemHeight = CType(item, Control).ActualHeight
'            measureItem.itemData = IntPtr.Zero
'            Dim msg As MSG
'            msg.hwnd = hwnd ' Handle to the window that owns the menu
'            msg.message = WM.MEASUREITEM
'            msg.wParam = IntPtr.Zero
'            msg.lParam = Marshal.AllocHGlobal(Marshal.SizeOf(measureItem))
'            Marshal.StructureToPtr(measureItem, msg.lParam, False)

'            contextMenu2.HandleMenuMsg(msg.message, msg.wParam, msg.lParam)

'            CType(item, Control).Width = measureItem.itemWidth
'            CType(item, Control).Height = measureItem.itemHeight

'            Dim hdc As IntPtr = Functions.GetDC(hwnd)
'            Dim dis As New DRAWITEMSTRUCT()
'            dis.CtlType = ODT.ODT_MENU ' Indicating it's a menu item
'            dis.CtlID = 0 ' Not used for menu items
'            dis.itemID = item.Tag.ToString().Split(vbTab)(0) ' The ID of the menu item
'            dis.itemAction = ODA.ODA_DRAWENTIRE ' The action to be performed
'            dis.itemState = ODS.ODS_SELECTED ' The state of the item
'            dis.hwndItem = hMenu ' Handle to the menu
'            dis.hDC = hdc ' Handle to the device context
'            dis.rcItem = New Rect(0, 0, measureItem.itemWidth, measureItem.itemHeight)
'            dis.itemData = IntPtr.Zero ' Application-defined value

'            Marshal.FreeHGlobal(msg.lParam)

'            msg = New MSG()
'            msg.hwnd = hwnd ' Handle to the window that owns the menu
'            msg.message = WM.DRAWITEM
'            msg.wParam = IntPtr.Zero
'            msg.lParam = Marshal.AllocHGlobal(Marshal.SizeOf(dis))
'            Marshal.StructureToPtr(dis, msg.lParam, False)

'            ' Send the message
'            contextMenu2.HandleMenuMsg(msg.message, msg.wParam, msg.lParam)
'            'Functions.SendMessage(msg.hwnd, msg.message, msg.wParam, msg.lParam)

'            ' Clean up
'            Marshal.FreeHGlobal(msg.lParam)
'        Next
'    End Function)