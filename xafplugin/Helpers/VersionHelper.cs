using Microsoft.Win32;
using System;
using System.Deployment.Application;
using System.Diagnostics;
using System.Globalization;
using System.Reflection;

namespace xafplugin.Helpers
{
    public static class VersionHelper
    {
        // Cache de berekende versie zodat we niet elke keer de registry lezen
        private static string _cached;

        private const string AddinName = "xafinsight"; // pas aan als je FriendlyName wijzigt
        private const string DeploymentUrlMarker = "/xafinsight/clickonce/"; // jouw Pages-pad

        public static string GetDisplayVersion()
        {
            if (!string.IsNullOrEmpty(_cached))
                return _cached;

            // 1) ClickOnce publish version (bv. 1.0.0.37)
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                {
                    _cached = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
                    return _cached;
                }
            }
            catch { /* ignore */ }

            // 2) VSTO registry (ManifestVersion)
            try
            {
                const string keyPath = @"Software\Microsoft\VSTO\SolutionMetadata";
                using (var key = Registry.CurrentUser.OpenSubKey(keyPath))
                {
                    if (key != null)
                    {
                        foreach (var name in key.GetSubKeyNames())
                        {
                            using (var sub = key.OpenSubKey(name))
                            {
                                var friendly = (sub?.GetValue("FriendlyName") as string)?.ToLowerInvariant();
                                var url = (sub?.GetValue("DeploymentManifestUrl") as string)?.ToLowerInvariant();
                                var manifestV = sub?.GetValue("ManifestVersion") as string;

                                if (!string.IsNullOrEmpty(manifestV) &&
                                    ((friendly != null && friendly.Contains(AddinName)) ||
                                     (url != null && url.Contains(DeploymentUrlMarker))))
                                {
                                    _cached = manifestV;
                                    return _cached;
                                }
                            }
                        }
                    }
                }
            }
            catch { /* ignore */ }

            // 3) Fallback: assembly file/product version
            try
            {
                var path = Assembly.GetExecutingAssembly().Location;
                var fvi = FileVersionInfo.GetVersionInfo(path);
                _cached = string.IsNullOrWhiteSpace(fvi.ProductVersion) ? fvi.FileVersion : fvi.ProductVersion;
                return _cached ?? "Unknown";
            }
            catch { /* ignore */ }

            return _cached = "Unknown";
        }

        public static void ShowVersionPopup()
        {
            try
            {
                System.Windows.Forms.MessageBox.Show(
                    $"XafInsight Add-in{Environment.NewLine}{Environment.NewLine}Versie: {GetDisplayVersion()}",
                    "XafInsight - Versie",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Information
                );
            }
            catch
            {
                // geen WinForms? Laat in log zien als fallback
                Debug.WriteLine("XafInsight Version: " + GetDisplayVersion(), CultureInfo.InvariantCulture);
            }
        }
    }
}
