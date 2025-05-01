using Windows.ApplicationModel;
using Windows.Management.Deployment;

namespace Laila.Shell.WinRT
{
    public sealed class PackageHelper
    {
        public static bool IsRunningPackaged()
        {
            try
            {
                var dummy = Package.Current;
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        private static Dictionary<string, string> _lastPackages = new Dictionary<string, string>();

        public static bool IsPackagesUpdated()
        {
            // get current packages
            var packageManager = new PackageManager();
            var packages = packageManager.FindPackagesForUser("");
            Dictionary<string, string> currentPackages = new Dictionary<string, string>();
            foreach (var package in packages)
                currentPackages.Add(package.Id.FullName, $"{package.Id.Version.Major}.{package.Id.Version.Minor}.{package.Id.Version.Build}.{package.Id.Version.Revision}");

            try
            {
                // any new or updated packages?
                foreach (var package in currentPackages)
                {
                    if (!_lastPackages.ContainsKey(package.Key))
                        return true;
                    if (_lastPackages[package.Key] != package.Value)
                        return true;
                }

                // any removed packages?
                foreach (var package in _lastPackages)
                {
                    if (!currentPackages.ContainsKey(package.Key))
                        return true;
                }

                return false;
            }
            finally
            {
                // update last packages
                _lastPackages = currentPackages;
            }
        }
    }
}