using System;
using xafplugin.Modules;

namespace xafplugin.Interfaces
{
    public interface ISettingsProvider
    {
        FileSettings Get(string fileKey);
        void Set(string fileKey, Action<FileSettings> updateAction);
        void Save(string fileKey, FileSettings settings);
        void Reset(string fileKey);
        void ResetAll();
    }

}
