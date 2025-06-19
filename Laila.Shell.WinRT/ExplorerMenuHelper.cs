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
using Windows.ApplicationModel.Resources;
using Windows.ApplicationModel.Resources.Core;
using Windows.ApplicationModel;

namespace Laila.Shell.WinRT
{
    public sealed class ExplorerMenuHelper
    {
        private static List<ExplorerCommandVerbInfo>? _cachedFileExplorerVerbs = null;
            
        public async Task<Tuple<string?, string?, List<ExplorerCommandVerbInfo>?>?> GetCloudVerbs(string? folderPath)
        {
            try
            {
                if (folderPath == null || !System.IO.Directory.Exists(folderPath))
                    return null; // cannot get verbs for files in multiple different folders

                // get sync root info
                string? applicationId = await new StorageHelper().IsCloudFolder(folderPath);
                if (applicationId == null) return null; // not a cloud folder

                // find appx manifest
                Tuple<string, string?, string, string?, string?>? application = findApplication(applicationId);
                if (application == null) return null;

                // get verbs
                List<ExplorerCommandVerbInfo>? explorerCommandVerbs = parseCloudVerbs(application.Item1, application.Item2);
                if (explorerCommandVerbs != null)
                    foreach (var verb in explorerCommandVerbs)
                    {
                        if (verb.ComServerPath != null && !System.IO.File.Exists(verb.ComServerPath))
                            if (System.IO.File.Exists(System.IO.Path.Combine(application.Item3, verb.ComServerPath)))
                                verb.ComServerPath = System.IO.Path.Combine(application.Item3, verb.ComServerPath);
                            else if (!string.IsNullOrWhiteSpace(System.IO.Path.GetDirectoryName(application.Item1)) 
                                && System.IO.File.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(application.Item1)!, verb.ComServerPath)))
                                    verb.ComServerPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(application.Item1)!, verb.ComServerPath);
                    }

                return new Tuple<string?, string?, List<ExplorerCommandVerbInfo>?>(
                    application.Item4,
                    application.Item5,
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

        public Task<List<ExplorerCommandVerbInfo>> GetFileExplorerVerbs()
        {
            if (PackageHelper.IsPackagesUpdated() || _cachedFileExplorerVerbs == null)
            {
                var verbs = new List<ExplorerCommandVerbInfo>();

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

                                foreach (XElement app in appElements)
                                {
                                    // get application
                                    Tuple<string, string?, string, string?, string?>? application = getApplication(package, manifestPath, app.Attribute("Id"), app);

                                    // get verbs
                                    List<ExplorerCommandVerbInfo>? explorerCommandVerbs = parseFileExplorerVerbs(application.Item1, application.Item2);

                                    // add items
                                    if (explorerCommandVerbs != null && explorerCommandVerbs.Count > 0)
                                    {
                                        foreach (var verb in explorerCommandVerbs)
                                        {
                                            if (verb.ComServerPath != null && !System.IO.File.Exists(verb.ComServerPath))
                                                if (System.IO.File.Exists(System.IO.Path.Combine(application.Item3, verb.ComServerPath)))
                                                    verb.ComServerPath = System.IO.Path.Combine(application.Item3, verb.ComServerPath);
                                                else if (!string.IsNullOrWhiteSpace(System.IO.Path.GetDirectoryName(application.Item1))
                                                    && System.IO.File.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(application.Item1)!, verb.ComServerPath)))
                                                    verb.ComServerPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(application.Item1)!, verb.ComServerPath);
                                            verb.PackageId = package.Id.Name;
                                            verb.ApplicationName = application.Item4;
                                            verb.ApplicationIconPath = application.Item5;
                                            verb.ApplicationExecutable = app?.Attribute("Executable")?.Value;
                                            verb.ManifestPath = manifestPath;
                                            verb.InstalledPath = package.EffectivePath;
                                        }
                                        verbs.AddRange(explorerCommandVerbs);
                                    }
                                }
                            }
                        }
                    }
                    catch
                    {
                        // Some packages are protected (Access Denied), skip them
                    }
                }

                _cachedFileExplorerVerbs = verbs;
                return Task.FromResult(verbs);
            }
            else
                return Task.FromResult(_cachedFileExplorerVerbs);
        }

        static Tuple<string, string?, string, string?, string?>? findApplication(string applicationId)
        {
            PackageManager packageManager = new PackageManager();
            var packages = packageManager.FindPackagesForUser("");

            foreach (Package package in packages)
            {
                try
                {
                    string manifestPath = System.IO.Path.Combine(package.InstalledLocation.Path, "AppxManifest.xml");

                    if (System.IO.File.Exists(manifestPath))
                    {
                        XDocument doc = XDocument.Load(manifestPath);
                        if (doc != null && doc.Root != null)
                        {
                            XNamespace ns = doc.Root.Name.Namespace;
                            var appElements = doc.Descendants(ns + "Application");

                            foreach (XElement app in appElements)
                            {
                                var idAttr = app.Attribute("Id");
                                if (idAttr != null && idAttr.Value.Equals(applicationId, StringComparison.OrdinalIgnoreCase))
                                    return getApplication(package, manifestPath, idAttr, app);
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

        static Tuple<string, string?, string, string?, string?> getApplication(Package package, string manifestPath, XAttribute? idAttr, XElement app)
        {
            string? displayName = package.DisplayName;
            if (displayName.StartsWith("Microsoft "))
                displayName = displayName.Substring("Microsoft ".Length);
            if (displayName.StartsWith("Windows "))
                displayName = displayName.Substring("Windows ".Length);
            string? logoPath = package.Logo.LocalPath;
            return new Tuple<string, string?, string, string?, string?>(manifestPath, idAttr?.Value, package.EffectivePath, displayName, logoPath);
        }

        /// <summary>
        /// Resolves an ms-resource key from a specific AppX package.
        /// </summary>
        /// <param name="resourceReference">The ms-resource reference (e.g., "ms-resource:MyDisplayName").</param>
        /// <param name="packageFamilyName">The full PackageFamilyName to load from (e.g., "DropboxInc.Dropbox_abcdefg123456").</param>
        /// <returns>The resolved localized string, or null if not found.</returns>
        public static string ResolveMsResourceFromPackage(string resourceReference, string packageFullName, string packageName)
        {
            const string prefix = "ms-resource:";
            string result = resourceReference;

            if (resourceReference.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                try
                {
                    if (packageName == null)
                        packageName = packageFullName.Split('_')[0];

                    List<IStorageFile> priFiles = new List<IStorageFile>() { StorageFile.GetFileFromPathAsync(ResourceLoader.GetDefaultPriPath(packageFullName)).Get() };

                    var resourceManager = ResourceManager.Current;
                    resourceManager.LoadPriFiles(priFiles);

                    var candidate = resourceManager.AllResourceMaps[packageName]?.GetValue(resourceReference);
                    if (candidate != null && candidate.IsMatch)
                        result = candidate.ValueAsString;
                    else
                    {
                       candidate = resourceManager.AllResourceMaps[packageName]?.GetValue(resourceManager.AllResourceMaps[packageName]?.Uri.AbsoluteUri + "Resources/" + resourceReference.Substring(prefix.Length));
                        if (candidate != null && candidate.IsMatch)
                            result = candidate.ValueAsString;
                    }

                    //resourceManager.UnloadPriFiles(priFiles);
                }
                catch
                {
                    // Ignore errors - resource not found or context issue
                }
            }

            // Not a ms-resource: string, return as-is
            return result;
        }

        static List<ExplorerCommandVerbInfo>? parseCloudVerbs(string manifestPath, string? applicationId)
        {
            var verbs = new List<ExplorerCommandVerbInfo>();

            try
            {
                XDocument doc = XDocument.Load(manifestPath);
                if (doc == null || doc.Root == null) return null;

                XNamespace ns = "http://schemas.microsoft.com/appx/manifest/foundation/windows10";
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
                var verbElements = doc.Descendants(ns + "Application").FirstOrDefault(app => app.Attribute("Id")?.Value == applicationId)?
                                      .Descendants(nsDesktop3 + "CloudFiles")
                                      .Descendants(nsDesktop3 + "Verb");

                if (verbElements != null)
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

        static List<ExplorerCommandVerbInfo>? parseFileExplorerVerbs(string manifestPath, string? applicationId)
        {
            var verbs = new List<ExplorerCommandVerbInfo>();

            try
            {
                XDocument doc = XDocument.Load(manifestPath);
                if (doc == null || doc.Root == null) return null;

                XNamespace ns = "http://schemas.microsoft.com/appx/manifest/foundation/windows10";
                XNamespace nsAppx = doc.Root.Name.Namespace;
                XNamespace nsDesktop4 = "http://schemas.microsoft.com/appx/manifest/desktop/windows10/4";
                XNamespace nsDesktop5 = "http://schemas.microsoft.com/appx/manifest/desktop/windows10/5";
                XNamespace nsDesktop10 = "http://schemas.microsoft.com/appx/manifest/desktop/windows10/10";
                XNamespace nsCom = "http://schemas.microsoft.com/appx/manifest/com/windows10";

                // 1. Find all COM classes
                var comClasses = doc.Descendants(nsCom + "Class")
                                    .ToDictionary(
                                        cls => ((string)cls.Attribute("Id")!).Trim('{', '}').ToUpperInvariant(),
                                        cls => (string)cls.Attribute("Path")!
                                    );

                // 2. Find all CloudFiles verbs
                var itemTypeElements = doc.Descendants(ns + "Application").FirstOrDefault(app => app.Attribute("Id")?.Value == applicationId)?
                                          .Descendants(nsDesktop4 + "FileExplorerContextMenus")
                                          .Descendants(nsDesktop4 + "ItemType")
                                          .Union(
                                              doc.Descendants(ns + "Application").FirstOrDefault(app => app.Attribute("Id")?.Value == applicationId)?
                                                 .Descendants(nsDesktop4 + "FileExplorerContextMenus")
                                                 .Descendants(nsDesktop5 + "ItemType")!
                                          )
                                          .Union(
                                              doc.Descendants(ns + "Application").FirstOrDefault(app => app.Attribute("Id")?.Value == applicationId)?
                                                 .Descendants(nsDesktop4 + "FileExplorerContextMenus")
                                                 .Descendants(nsDesktop10 + "ItemType")!
                                          );

                if (itemTypeElements != null)
                    foreach (var itemTypeElement in itemTypeElements)
                    {
                        var typeAttr = itemTypeElement.Attribute("Type");

                        if (typeAttr != null)
                        {
                            var itemType = typeAttr.Value;
                            var verbElements = itemTypeElement.Descendants(nsDesktop4 + "Verb")
                                        .Union(itemTypeElement.Descendants(nsDesktop5 + "Verb"))
                                        .Union(itemTypeElement.Descendants(nsDesktop10 + "Verb"));

                            if (verbElements != null)
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
                                            ComServerPath = serverPath,
                                            ItemType = itemType,
                                            ApplicationId = applicationId
                                        });
                                    }
                                }
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
