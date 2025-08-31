using System;

namespace xafplugin.Modules
{
    public enum EFileType
    {
        Csv,
        Excel,
        Json,
        Xml,
        db
    }

    public static class EFileTypeExtensions
    {
        public static string GetExtension(this EFileType type)
        {
            return "." + type.ToString();
        }

        public static string GetExtensionNoDot(this EFileType type)
        {
            return type.ToString();
        }
    }
}