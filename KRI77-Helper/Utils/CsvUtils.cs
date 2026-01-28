using System;
using System.Collections.Generic;
using System.Text;

namespace KRI77_Helper.Utils
{
    public static class CsvUtils
    {
        public static string[] SplitCsvLine(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
                return Array.Empty<string>();

            var result = new List<string>();
            var current = new StringBuilder();
            bool inQuotes = false;

            for (int i = 0; i < line.Length; i++)
            {
                char c = line[i];

                if (c == '"')
                {
                    // Handle escaped quotes ("")
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++; // skip next quote
                    }
                    else
                    {
                        inQuotes = !inQuotes;
                    }
                }
                else if (c == ',' && !inQuotes)
                {
                    result.Add(current.ToString());
                    current.Clear();
                }
                else
                {
                    current.Append(c);
                }
            }

            // Add last value
            result.Add(current.ToString());

            return result.ToArray();
        }


        public static class CsvLogger
        {
            /********************************************************************/
            //
            // Old
            // Console.WriteLine("Process started");
            //
            // New
            // CsvLogger.Log("Process started");
            //
            // With levels
            // CsvLogger.Log("Something unexpected happened", level: "WARN");
            // CsvLogger.Log("An error occurred while saving record 123", level: "ERROR");
            /********************************************************************/

            private static readonly object _lock = new object();
            private static readonly string _logPath = Path.Combine(
                AppContext.BaseDirectory, "logs", $"app-log-{DateTime.UtcNow:yyyyMMdd}.csv");

            private static bool _headerWritten = false;

            public static void Log(string message, string level = "INFO")
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_logPath)!);

                lock (_lock)
                {
                    var sb = new StringBuilder();

                    // Write CSV header once per file
                    if (!_headerWritten && (!File.Exists(_logPath) || new FileInfo(_logPath).Length == 0))
                    {
                        sb.AppendLine("TimestampUTC,Level,Message");
                        _headerWritten = true;
                    }

                    // Escape commas/quotes/newlines properly for CSV
                    string Escape(string s) =>
                        "\"" + (s ?? string.Empty).Replace("\"", "\"\"") + "\"";

                    sb.AppendLine($"{DateTime.UtcNow:O},{Escape(level)},{Escape(message)}");
                    File.AppendAllText(_logPath, sb.ToString(), Encoding.UTF8);
                }
            }

            public static void Log(string input, string output, string archive, int rowsProcessed, int rowsDuplicate)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(_logPath)!);

                lock (_lock)
                {
                    var sb = new StringBuilder();

                    // Write CSV header once per file
                    if (!_headerWritten && (!File.Exists(_logPath) || new FileInfo(_logPath).Length == 0))
                    {
                        sb.AppendLine("Date,Time,Input,Output,Archive,Rows,Duplicates");
                        _headerWritten = true;
                    }

                    // Escape commas/quotes/newlines properly for CSV
                    string Escape(string s) =>
                        "\"" + (s ?? string.Empty).Replace("\"", "\"\"") + "\"";

                    string dt = DateTime.UtcNow.ToString("yyyy-MM-dd");
                    string tm = DateTime.UtcNow.ToString("hh:mm tt");

                    sb.AppendLine($"{dt},{tm},{Escape(input)},{Escape(output)},{Escape(archive)},{Escape(rowsProcessed.ToString())},{Escape(rowsDuplicate.ToString())}");
                    File.AppendAllText(_logPath, sb.ToString(), Encoding.UTF8);
                }
            }
        }

    }
}