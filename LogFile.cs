using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PDF_Report_Generator.Logging
{
    class LogFile_PDFGenerator
    {
        public static string GetLogFile()
        {
            string logFilePath;
            string destPath = ConfigurationManager.AppSettings["destFilePath"].ToString();
            string currentLogFile = "PDFGenerator_Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            logFilePath = destPath + "\\Logs\\" + currentLogFile;

            if (!File.Exists(logFilePath))
            {
                // Create the file.
                using (FileStream fs = File.Create(logFilePath))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes("===============================PDF Generator Log File " + DateTime.Now.ToShortDateString() + "===================================" + Environment.NewLine);

                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);

                }

            }
            return logFilePath;

        }

        public static void WriteLogMessage(string logMessage)
        {
            string logFilePath = GetLogFile();

            using (StreamWriter w = File.AppendText(logFilePath))
            {
                Log(logMessage, w);
            }

            //using (StreamReader r = File.OpenText(logFilePath))
            //{
            //    DumpLog(r);
            //}
        }

        public static void WriteBlankLine()
        {
            string logFilePath = GetLogFile();

            using (StreamWriter w = File.AppendText(logFilePath))
            {
                w.WriteLine("      ");
            }
        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("{0} ", DateTime.Now.ToLongTimeString());
            w.WriteLine("  {0}", logMessage);

        }

        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
    }

    class LogFile_CustomerWeekly
    {
        public static string GetLogFile()
        {
            string logFilePath;
            string destPath = ConfigurationManager.AppSettings["destFilePath"].ToString();
            string currentLogFile = "CustomerWeeklyPDF_Log_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            logFilePath = destPath + "\\Logs\\" + currentLogFile;
            
            if (!File.Exists(logFilePath))
            {
                // Create the file.
                using (FileStream fs = File.Create(logFilePath))
                {
                    Byte[] info = new UTF8Encoding(true).GetBytes("===============================Customer Weekly Generate PDF Log File " + DateTime.Now.ToShortDateString() + "===================================" + Environment.NewLine);

                    // Add some information to the file.
                    fs.Write(info, 0, info.Length);
                    
                }

            }
            return logFilePath;

        }

        public static void WriteLogMessage(string logMessage)
        {
            string logFilePath = GetLogFile();

            using (StreamWriter w = File.AppendText(logFilePath))
            {
                Log(logMessage, w);
            }

            //using (StreamReader r = File.OpenText(logFilePath))
            //{
            //    DumpLog(r);
            //}
        }

        public static void WriteBlankLine()
        {
            string logFilePath = GetLogFile();

            using (StreamWriter w = File.AppendText(logFilePath))
            {
                w.WriteLine("      ");
            }
        }

        public static void Log(string logMessage, TextWriter w)
        {
            w.Write("{0} ", DateTime.Now.ToLongTimeString());
            w.WriteLine("  {0}", logMessage);

        }

        public static void DumpLog(StreamReader r)
        {
            string line;
            while ((line = r.ReadLine()) != null)
            {
                Console.WriteLine(line);
            }
        }
    }
}
