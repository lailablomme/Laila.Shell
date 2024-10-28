using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.DataTransfer;
using Windows.Foundation;
using Windows.Storage.Streams;
using Windows.Storage;
using System.Collections;
using static Laila.Shell.WinRT.DataTransferManagerHelper;

namespace Laila.Shell.WinRT
{
    static class DataTransferManagerHelper
    {
        static readonly Guid _dtm_iid = new Guid(0xa5caee9b, 0x8708, 0x49d1, 0x8d, 0x36, 0x67, 0xd2, 0x5a, 0x8d, 0xa0, 0x0c);

        static IDataTransferManagerInterop DataTransferManagerInterop
        {
            get
            {
                return DataTransferManager.As<IDataTransferManagerInterop>();
            }
        }

        public static DataTransferManager GetForWindow(IntPtr hwnd)
        {
            return DataTransferManager.FromAbi(DataTransferManagerInterop.GetForWindow(hwnd, _dtm_iid));
        }

        public static void ShowShareUIForWindow(IntPtr hwnd)
        {
            DataTransferManagerInterop.ShowShareUIForWindow(hwnd);
        }

        [ComImport, Guid("3A3DCD6C-3EAB-43DC-BCDE-45671CE800C8")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IDataTransferManagerInterop
        {
            IntPtr GetForWindow([In] IntPtr appWindow, [In] ref Guid riid);
            void ShowShareUIForWindow(IntPtr appWindow);
        }

        public delegate void DataRequestedEventHandler(IDataTransferManager sender, IDataRequestedEventArgs args);
        public delegate void TargetApplicationChosenEventHandler(IDataTransferManager sender, ITargetApplicationChosenEventArgs args); [ComImport, Guid("A5CAEE9B-8708-49D1-8D36-67D25A8DA00C")]
        public interface IDataTransferManager
        {
            event DataRequestedEventHandler DataRequested;
            event TargetApplicationChosenEventHandler TargetApplicationChosen;
        }

        [ComImport, Guid("CB8BA807-6AC5-43C9-8AC5-9BA232163182")]
        internal interface IDataRequestedEventArgs
        {
            IDataRequest Request { get; }
        }

        [ComImport, Guid("CA6FB8AC-2987-4EE3-9C54-D8AFBCB86C1D")]
        internal interface ITargetApplicationChosenEventArgs
        {
            string ApplicationName { get; }
        }
        
        [ComImport, Guid("4341AE3B-FC12-4E53-8C02-AC714C415A27")]
        internal interface IDataRequest
        {
            IDataPackage Data { get; set; }

            //DateTimeOffset Deadline { get; }

            //void FailWithDisplayText(string value);

            //DataRequestDeferral GetDeferral();
        }

        [ComImport, Guid("61EBF5C7-EFEA-4346-9554-981D7E198FFE")]
        internal interface IDataPackage
        {
            IDataPackagePropertySet Properties { get; }

            //DataPackageOperation RequestedOperation { get; set; }

            //IDictionary<string, RandomAccessStreamReference> ResourceMap { get; }

            //event TypedEventHandler<DataPackage, object> Destroyed;

            //event TypedEventHandler<DataPackage, OperationCompletedEventArgs> OperationCompleted;

            //DataPackageView GetView();

            //void SetData(string formatId, object value);

            //void SetDataProvider(string formatId, DataProviderHandler delayRenderer);

            //void SetText(string value);

            //void SetUri(Uri value);

            //void SetHtmlFormat(string value);

            //void SetRtf(string value);

            //void SetBitmap(RandomAccessStreamReference value);

            //[Windows.Foundation.Metadata.Overload("SetStorageItemsReadOnly")]
            void SetStorageItems(IEnumerable<IStorageItem> value);

            //[Overload("SetStorageItems")]
            //void SetStorageItems(IEnumerable<IStorageItem> value, bool readOnly);
        }

        [ComImport, Guid("CD1C93EB-4C4C-443A-A8D3-F5C241E91689")]
        internal interface IDataPackagePropertySet //: IDictionary<string, object>, ICollection<KeyValuePair<string, object>>, IEnumerable<KeyValuePair<string, object>>, IEnumerable
        {
            Uri ApplicationListingUri { get; set; }

            string ApplicationName { get; set; }

            string Description { get; set; }

            IList<string> FileTypes { get; }

            IRandomAccessStreamReference Thumbnail { get; set; }

            string Title { get; set; }
        }
    }
}
