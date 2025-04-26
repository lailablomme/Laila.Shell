using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Windows.Storage.Provider;
using Windows.Storage;
using Laila.Shell.WinRT.Interface;
using System.Diagnostics;
using System.Xml.Linq;
using Windows.Management.Deployment;

namespace Laila.Shell.WinRT
{
    public class ExplorerMenuHelper
    {
        public async Task<Tuple<string, string, List<ExplorerCommandVerbInfo>?>?> GetCloudVerbs(string? folderPath)
        {
            try
            {
                if (folderPath == null)
                    return null; // cannot get verbs for files in multiple different folders

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
                }
                else
                {
                    Debug.WriteLine("Not a cloud folder.");
                    return null;
                }

                // find appx manifest
                Tuple<string, string>? paths = findAppxManifest(applicationId);
                if (paths == null) return null;

                // get verbs
                List<ExplorerCommandVerbInfo>? explorerCommandVerbs = ParseCloudVerbs(paths.Item1);
                if (explorerCommandVerbs != null)
                    foreach (var verb in explorerCommandVerbs)
                    {
                        if (verb.ComServerPath != null && !System.IO.File.Exists(verb.ComServerPath))
                            if (System.IO.File.Exists(System.IO.Path.Combine(paths.Item2, verb.ComServerPath)))
                                verb.ComServerPath = System.IO.Path.Combine(paths.Item2, verb.ComServerPath);
                            else if (!string.IsNullOrWhiteSpace(System.IO.Path.GetDirectoryName(paths.Item1)) 
                                && System.IO.File.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(paths.Item1)!, verb.ComServerPath)))
                                    verb.ComServerPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(paths.Item1)!, verb.ComServerPath);
                    }

                return new Tuple<string, string, List<ExplorerCommandVerbInfo>?>(
                    syncInfo.DisplayNameResource,
                    syncInfo.IconResource,
                    explorerCommandVerbs
                );
            }
            catch (Exception ex)
            {
                // Handle exceptions
                Debug.WriteLine("Error: " + ex.Message);
                return null;
            }
        }

        static Tuple<string, string>? findAppxManifest(string applicationId)
        {
            var packageManager = new PackageManager();
            var packages = packageManager.FindPackagesForUser("");

            foreach (var package in packages)
            {
                try
                {
                    var manifestPath = System.IO.Path.Combine(package.InstalledLocation.Path, "AppxManifest.xml");

                    if (System.IO.File.Exists(manifestPath))
                    {
                        XDocument doc = XDocument.Load(manifestPath);
                        if (doc != null && doc.Root != null)
                        {
                            XNamespace ns = doc.Root.Name.Namespace;
                            var appElements = doc.Descendants(ns + "Application");

                            foreach (var app in appElements)
                            {
                                var idAttr = app.Attribute("Id");
                                if (idAttr != null && idAttr.Value.Equals(applicationId, StringComparison.OrdinalIgnoreCase))
                                    return new Tuple<string, string>(manifestPath, package.EffectivePath);
                            }
                        }
                    }
                }
                catch
                {
                    // Some packages are protected (Access Denied), skip them
                }
            }

            return null;
        }

        static List<ExplorerCommandVerbInfo>? ParseCloudVerbs(string manifestPath)
        {
            var verbs = new List<ExplorerCommandVerbInfo>();

            try
            {
                XDocument doc = XDocument.Load(manifestPath);
                if (doc == null || doc.Root == null) return null;

                XNamespace nsAppx = doc.Root.Name.Namespace;
                XNamespace nsDesktop3 = "http://schemas.microsoft.com/appx/manifest/desktop/windows10/3";
                XNamespace nsCom = "http://schemas.microsoft.com/appx/manifest/com/windows10";

                // 1. Find all COM classes
                var comClasses = doc.Descendants(nsCom + "Class")
                                    .ToDictionary(
                                        cls => ((string)cls.Attribute("Id")!).Trim('{', '}').ToUpperInvariant(),
                                        cls => (string)cls.Attribute("Path")!
                                    );

                // 2. Find all CloudFiles verbs
                var verbElements = doc.Descendants(nsDesktop3 + "CloudFiles")
                                      .Descendants(nsDesktop3 + "Verb");

                foreach (var verbElement in verbElements)
                {
                    var idAttr = verbElement.Attribute("Id");
                    var clsidAttr = verbElement.Attribute("Clsid");

                    if (idAttr != null && clsidAttr != null)
                    {
                        string verbId = idAttr.Value;
                        string clsid = clsidAttr.Value.Trim('{', '}').ToUpperInvariant();

                        string? serverPath = null;
                        comClasses.TryGetValue(clsid, out serverPath);

                        verbs.Add(new ExplorerCommandVerbInfo
                        {
                            Id = verbId,
                            ClsId = new Guid(clsid),
                            ComServerPath = serverPath
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed to parse manifest: " + ex.Message);
            }

            return verbs;
        }

    }
}
