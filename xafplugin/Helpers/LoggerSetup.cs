using NLog;
using NLog.Config;
using NLog.Targets;
using System;
using System.IO;

namespace xafplugin.Helpers
{
    public static class LoggerSetup
    {
        public static void Configure()
        {
            var config = new LoggingConfiguration();

            string logFolder = Globals.ThisAddIn.Config.LogPath;

            if (string.IsNullOrWhiteSpace(logFolder) || !Directory.Exists(logFolder))
            {
                throw new InvalidOperationException("LogPath is not configured.");
            }

            // Bestandslog target
            var logfile = new FileTarget("logfile")
            {
                FileName = Path.Combine(logFolder, "log.txt"),
                Layout = "${longdate}|${level:uppercase=true}|${_logger}|${message} ${exception:format=toString,stacktrace}",
                ArchiveEvery = FileArchivePeriod.Day,
                ArchiveFileName = Path.Combine(logFolder, "log.txt"),
                ArchiveSuffixFormat = "_yyyyMMdd",  // Bijv. log_20250718.txt
                MaxArchiveFiles = 7,
                KeepFileOpen = false,
                Encoding = System.Text.Encoding.UTF8
            };


            // Visual Studio Output venster target (alleen in debug)
#if DEBUG
            var debugOutput = new DebugTarget("debug")
            {
                Layout = "${longdate}|${level}|${_logger}|${message} ${exception:format=toString}"
            };
            config.AddTarget(debugOutput);
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, debugOutput);
#endif

            config.AddTarget(logfile);

#if DEBUG
            config.AddRule(LogLevel.Debug, LogLevel.Fatal, logfile);
#else
        config.AddRule(LogLevel.Info, LogLevel.Fatal, logfile);
#endif

            LogManager.Configuration = config;
            LogManager.ThrowExceptions = false;
        }
    }
}
