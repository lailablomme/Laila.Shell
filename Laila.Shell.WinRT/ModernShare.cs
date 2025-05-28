using Laila.Shell.WinRT.Properties;
using System.Windows;
using System.Windows.Interop;
using Windows.ApplicationModel.DataTransfer;
using Windows.Storage;

namespace Laila.Shell.WinRT
{
    public sealed class ModernShare
    {
        public void ShowShareUI(IEnumerable<string> filePaths, Window window)
        {
            IntPtr hwnd = new WindowInteropHelper(window).Handle;

            var files = new List<IStorageItem>();
            foreach (var filePath in filePaths)
            {
                var file = StorageFile.GetFileFromPathAsync(filePath).AsTask().Result;
                files.Add(file);
            }

            var dataTransferManager = DataTransferManagerHelper.GetForWindow(hwnd);
            dataTransferManager.DataRequested += (DataTransferManager sender, DataRequestedEventArgs args) =>
            {
                var requestData = args.Request.Data;
                requestData.Properties.Title = ResourcesRT.ModernShare_Title;
                requestData.Properties.Description = ResourcesRT.ModernShare_Description;
                requestData.SetStorageItems(files);
            };

            DataTransferManagerHelper.ShowShareUIForWindow(hwnd);
        }
    }
}



