using log4net;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;

namespace GorillaDocs
{
    public class Message
    {
        public static void Configure(FileInfo ConfigFile, string LogLevel)
        {
            log4net.Config.XmlConfigurator.Configure(ConfigFile);
            var repository = ((log4net.Repository.Hierarchy.Hierarchy)LogManager.GetRepository());
            repository.Root.Level = repository.LevelMap[LogLevel];
            repository.RaiseConfigurationChanged(EventArgs.Empty);
        }

        //See http://logging.apache.org/log4net/index.html
        public static void ShowError(Exception ex)
        {
            StackTrace stackTrace = new StackTrace();
            MethodBase method = stackTrace.GetFrame(1).GetMethod();
            MessageBox.Show(ex.Message, String.Format("{0} {1}", Assembly.GetExecutingAssembly().Title(), Assembly.GetExecutingAssembly().FileVersion()), MessageBoxButton.OK, MessageBoxImage.Exclamation);
            LogError(method, ex);
        }

        public static void ShowInformation(string message)
        {
            MessageBox.Show(message, String.Format("{0} {1}", Assembly.GetExecutingAssembly().Title(), Assembly.GetExecutingAssembly().FileVersion()), MessageBoxButton.OK, MessageBoxImage.Information);
            LogInfo(message);
        }

        public static void LogError(Exception ex)
        {
            StackTrace stackTrace = new StackTrace();
            MethodBase method = stackTrace.GetFrame(1).GetMethod();
            LogError(method, ex);
        }
        static void LogError(MethodBase method, Exception ex)
        {
            Task.Factory.StartNew(() =>
            {
                ILog log = LogManager.GetLogger(method.DeclaringType);
                log.Error(method.Name, ex);
            });
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static bool IsWarnEnabled()
        {
            ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            return log.IsWarnEnabled;
        }

        public static void LogWarning(Exception ex)
        {
            StackTrace stackTrace = new StackTrace();
            MethodBase method = stackTrace.GetFrame(1).GetMethod();
            LogWarning(method, string.Empty, ex);
        }
        public static void LogWarning(MethodBase method, Exception ex)
        {
            LogWarning(method, string.Empty, ex);
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void LogWarning(MethodBase method, string parameters, Exception ex)
        {
            Task.Factory.StartNew(() =>
            {
                ILog log = LogManager.GetLogger(method.DeclaringType);
                log.Warn(method.Name + parameters, ex);
            });
        }
        public static void LogWarning(string value)
        {
            Task.Factory.StartNew(() =>
            {
                StackTrace stackTrace = new StackTrace();
                MethodBase method = stackTrace.GetFrame(1).GetMethod();
                ILog log = LogManager.GetLogger(method.DeclaringType);
                log.Warn(value);
            });
        }
        public static void LogWarning(string format, params object[] args)
        {
            LogWarning(string.Format(format, args));
        }

        public static void LogInfo(string value)
        {
            Task.Factory.StartNew(() =>
            {
                StackTrace stackTrace = new StackTrace();
                MethodBase method = stackTrace.GetFrame(1).GetMethod();
                ILog log = LogManager.GetLogger(method.DeclaringType);
                log.InfoFormat("{0} {1}", method.Name, value);
            });
        }
        public static void LogInfo(string format, params object[] args)
        {
            LogInfo(string.Format(format, args));
        }

        [System.Diagnostics.DebuggerStepThrough]
        public static bool IsDebugEnabled()
        {
            ILog log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
            return log.IsDebugEnabled;
        }

        public static void LogDebug(string value)
        {
            StackTrace stackTrace = new StackTrace();
            MethodBase method = stackTrace.GetFrame(1).GetMethod();
            LogDebug(method, value);
        }
        [System.Diagnostics.DebuggerStepThrough]
        public static void LogDebug(MethodBase method, string value)
        {
            Task.Factory.StartNew(() =>
            {
                ILog log = LogManager.GetLogger(method.DeclaringType);
                if (log.IsDebugEnabled) log.DebugFormat("{0} {1}", method.Name, value);
            });
        }
    }
}
