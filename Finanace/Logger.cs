using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml.Linq;

namespace FinanceApplication
{
    public class Logger
    {
        private string Filename = String.Format("FinanceLogger_{0:yyyy_mm_dd__HH_mm_ss}.txt",  DateTime.Now);
        private string PathToLog = System.Configuration.ConfigurationManager.AppSettings["LogLocation"];
        private StreamWriter stream;
        private static Logger _instance;

        private Logger(bool CleanLog)
        {
            Initialize(CleanLog);
        }

        private Logger()
        {
            Initialize(false);
        }

        public static Logger CreateLogger(bool CleanLog)
        {
            if (_instance == null)
            {
                _instance = new Logger(CleanLog);
            }
            return _instance;
        }

        public static Logger CreateLogger()
        {
            return CreateLogger(false);
        }

        private void Initialize(bool clean)
        {
            string fullpath = Path.Combine(PathToLog, Filename);
            if (!Directory.Exists(PathToLog))
            {
                Directory.CreateDirectory(PathToLog);
            }

            if (clean && File.Exists(fullpath))
            {
                File.Delete(fullpath);
            }

            stream = new StreamWriter(File.OpenWrite(fullpath));
        }

        public void WriteError(string format, params object[] arg0)
        {
            Console.WriteLine(String.Format("ERROR: {0}", String.Format(format, arg0)));
            stream.WriteLine(String.Format("ERROR: {0}", String.Format(format, arg0)));
        }

        public void WriteWarning(string format, params object[] arg0)
        {
            Console.WriteLine(String.Format("Warning: {0}", String.Format(format, arg0)));
            stream.WriteLine(String.Format("Warning: {0}", String.Format(format, arg0)));
        }
        
        public void WriteInfo(string format, params object[] arg0)
        {
            Console.WriteLine(String.Format(format, arg0));
            stream.WriteLine(String.Format(format, arg0));
        }

        public void Close()
        {
            stream.Flush();
            stream.Close();
        }

    }
}
