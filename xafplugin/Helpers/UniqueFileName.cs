using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using xafplugin.Modules;

namespace xafplugin.Helpers
{
    public static class UniqueFileName
    {
        public static string GetFileHash(string filePath)
        {
            try
            {
                using (var sha256 = SHA256.Create())
                using (var stream = File.OpenRead(filePath))
                {
                    byte[] hash = sha256.ComputeHash(stream);
                    return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                }
            }
            catch (Exception ex)
            {
                return GetHashFromString(filePath);
            }
        }

        public static string GetFileHash(params string[] filePaths)
        {
            if (filePaths == null || filePaths.Length == 0)
                return string.Empty;

            if (filePaths.Length == 1)
                return GetFileHash(filePaths[0]);

            var fileHashes = filePaths
                .Where(File.Exists)
                .Select(path =>
                {
                    long len = new FileInfo(path).Length;
                    return BitConverter.GetBytes(len);
                })
                .OrderBy(b => BitConverter.ToUInt64(b, 0)) // sort for order-independence
                .ToArray();

            using (var sha256 = SHA256.Create())
            {
                foreach (var fh in fileHashes)
                {
                    sha256.TransformBlock(fh, 0, fh.Length, null, 0);
                }
                sha256.TransformFinalBlock(new byte[0], 0, 0);

                return BytesToHex(sha256.Hash);
            }
        }

        private static string BytesToHex(byte[] bytes)
        {
            // Equivalent to Convert.ToHexString in .NET 5+
            return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
        }

        private static string GetHashFromString(string input)
        {
            using (var sha256 = SHA256.Create())
            {
                byte[] hash = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }


        public static string CombinePathAndName(string fileKey, string path, EFileType extention)
        {
            return Path.Combine(path, fileKey + extention.GetExtension());
        }
    }
}
