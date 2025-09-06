using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace xafplugin.Helpers
{
    public static class VersionHelper
    {
        public static string GetDisplayVersion()
        {
            // 1) ClickOnce publish version (bv. 1.0.0.37)
            try
            {
                if (ApplicationDeployment.IsNetworkDeployed)
                    return ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            }
            catch { /* ignore */ }

            // 2) Fallback: file/product version uit assembly
            try
            {
                var path = Assembly.GetExecutingAssembly().Location;
                var fvi = FileVersionInfo.GetVersionInfo(path);
                return string.IsNullOrWhiteSpace(fvi.ProductVersion) ? fvi.FileVersion : fvi.ProductVersion;
            }
            catch { /* ignore */ }

            return "Unknown";
        }
    }
}
