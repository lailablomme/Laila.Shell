using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.Storage.Provider;
using Windows.Storage;

namespace Laila.Shell.WinRT
{
    public sealed class StorageHelper
    {
        public async Task<string?> IsCloudFolder(string folderPath)
        {
            try
            {
                if (!System.IO.Directory.Exists(folderPath))
                {
                    Debug.WriteLine("Folder does not exist: " + folderPath);
                    return null;
                }

                // get sync root info
                StorageFolder folder = await StorageFolder.GetFolderFromPathAsync(folderPath);
                StorageProviderSyncRootInfo syncInfo = StorageProviderSyncRootManager.GetSyncRootInformationForFolder(folder);
                string? applicationId = null;
                if (syncInfo != null)
                {
                    if (string.IsNullOrWhiteSpace(syncInfo.Id))
                    {
                        Debug.WriteLine("Sync root ID is empty.");
                        return null;
                    }
                    if (syncInfo.Id.Contains("!"))
                        applicationId = syncInfo.Id.Split('!')[0];
                    else
                        applicationId = syncInfo.Id;

                    Console.WriteLine("✅ Found Sync Root Info!");
                    Console.WriteLine("Provider ID: " + syncInfo.Id);
                    Console.WriteLine("Display Name: " + syncInfo.DisplayNameResource);
                    Console.WriteLine("Icon Resource: " + syncInfo.IconResource);

                    return applicationId;
                }
                else
                {
                    Debug.WriteLine("Not a cloud folder.");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Error checking cloud folder: " + ex.Message);
                return null;
            }
        }
    }
}
