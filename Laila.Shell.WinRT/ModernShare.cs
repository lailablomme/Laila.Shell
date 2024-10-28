using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

namespace Laila.Shell.WinRT
{
    public class ModernShare
    {
        public void ShowShareUI(List<string> filePaths)
        {
            IntPtr hwnd = new WindowInteropHelper(System.Windows.Application.Current.MainWindow).Handle;

            var dataTransferManager = DataTransferManagerHelper.GetForWindow(hwnd);
            dataTransferManager.DataRequested += (DataTransferManager sender, DataRequestedEventArgs args) =>
            {
                var requestData = args.Request.Data;
                requestData.Properties.Title = "Share Files";
                requestData.Properties.Description = "Sharing multiple files.";

                var files = new List<IStorageItem>();
                foreach (var filePath in filePaths)
                {
                    var file = StorageFile.GetFileFromPathAsync(filePath).AsTask().Result;
                    files.Add(file);
                }

                requestData.SetStorageItems(files);
            };

            DataTransferManagerHelper.ShowShareUIForWindow(hwnd);
        }
    }
}



