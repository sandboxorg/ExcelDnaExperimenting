using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using System.Text;

namespace ExcelDna.Logging
{
    static class CustomLogging
    {
        static void EnsureInit()
        {
            if (_logger != null)
                return;

            var assemblies = AppDomain.CurrentDomain.GetAssemblies();
            Type loggerType = null;
            foreach (var assembly in assemblies)
            {
                foreach (var type in assembly.GetTypes())
                {
                    if (type.Name == "ExcelDnaLogTarget")
                    {
                        loggerType = type;
                        break;
                    }
                }
            }

            if (loggerType != null)
            {
                _logger = loggerType.GetMethod("Log", BindingFlags.Public | BindingFlags.Static);
            }
        }

        public static void Log(TraceEventType eventType, string message, params object[] args)
        {
            try
            {
                EnsureInit();

                var renderedMessage = string.Format(message, args);
                if (_logger != null)
                    _logger.Invoke(null, new object[] { eventType, renderedMessage });
            }
            catch (Exception exc)
            {
                Trace.WriteLine(exc);
            }
        }

        private static MethodInfo _logger;
    }
}
