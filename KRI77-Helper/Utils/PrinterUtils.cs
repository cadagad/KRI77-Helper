using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml.Office;
using KRI77_Helper.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KRI77_Helper.Utils.CsvUtils;

namespace KRI77_Helper.Utils
{
    internal class PrinterUtils
    {
        //public static List<Printer> ConsolidatePrintersAsia(string in_path)
        public static List<Printer> ConsolidatePrintersAsia(string in_path)
        {

            // Find immediate subfolders that start with "Asia Printers" (case-insensitive)
            IEnumerable<string> asiaFolders = Directory.EnumerateDirectories(in_path, "*", SearchOption.TopDirectoryOnly)
                .Where(d => Path.GetFileName(d).StartsWith("Asia Printers", StringComparison.OrdinalIgnoreCase));


            // From those folders, get ALL .csv / .xlsx files (including nested subfolders)
            IEnumerable<string> matchingFiles = asiaFolders
                .SelectMany(folder =>
                    Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                             .Where(f => f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase)
                                      || f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)));


            // (optional) materialize to a list
            List<string> files = matchingFiles.ToList();

            List<Printer> printer_asia = new List<Printer>();

            // Debug print
            foreach (var f in files)
            {
                string fileName = Path.GetFileName(f);

                if (fileName.Contains("China", StringComparison.OrdinalIgnoreCase))
                {
                    //Important: pass the FULL PATH to the processor unless it opens relative to in_path
                    var chinaPrinters = ProcessPrinterChina(f);   // or ProcessPrinterChina(in_path, f)
                    if (chinaPrinters != null && chinaPrinters.Count > 0)
                        printer_asia.AddRange(chinaPrinters);
                }
                else if (fileName.Contains("Japan", StringComparison.OrdinalIgnoreCase))
                {
                    
                    var japanPrinters = ProcessPrinterJapan(f);   // or ProcessPrinterChina(in_path, f)
                    if (japanPrinters != null && japanPrinters.Count > 0)
                        printer_asia.AddRange(japanPrinters);
                }
                else
                    continue;
            }

            //Console.WriteLine(printer_asia.Count);
            //Console.WriteLine($"Printer Asia Count: {printer_asia.Count}");
            CsvLogger.Log($"Printer Asia Count: {printer_asia.Count}");

            return printer_asia;
        }

        public static List<Printer> ProcessPrinterChina(string filepath)
        {
            int in_count_China = 0;
            //int out_count = 0;

            List<Printer> printer_china = new List<Printer>();

            /* Read Network NA Excel file and store to List <Network> - networks_na */
            // Console.WriteLine($"Processing file: China Printer Files");
            CsvLogger.Log($"Processing file: China Printer Files");
            using (var workbook = new XLWorkbook(Path.Combine(filepath)))
            {
                var worksheet = workbook.Worksheet(1); // Get first worksheet
                var rows = worksheet.RowsUsed();

                foreach (var row in rows)
                {
                    in_count_China++;

                    /* Skip header row (row 1) */
                    if (row.RowNumber() == 1)
                        continue;

                    /* SerialNumber cannot be empty */
                    string SerialNumber = row.Cell(1).GetValue<string>();
                    if (String.IsNullOrEmpty(SerialNumber) || SerialNumber == "N/A" || SerialNumber == "NA")
                        continue;

                    Printer printer = new Printer()
                    {
                        Country = "CHI",
                        Class = "Printer",
                        AssetTag = "",
                        SerialNumber = row.Cell(1).GetValue<string>(),
                        AssetStatus = row.Cell(5).GetValue<string>(),
                        Location = row.Cell(4).GetValue<string>(),
                        LocationDetail = "",
                        OwnedBy = "",
                        Model = row.Cell(3).GetValue<string>(),
                        SupportGroup = ""
                    };

                    printer_china.Add(printer);
                }
            }       
            
            return printer_china;
        }

        public static List<Printer> ProcessPrinterJapan(string filepath)
        {
            int in_count_China = 0;
            //int out_count = 0;

            List<Printer> printer_china = new List<Printer>();

            /* Read Network NA Excel file and store to List <Network> - networks_na */
            //Console.WriteLine($"Processing file: Japan Printer Files");
            CsvLogger.Log($"Processing file: Japan Printer Files");
            using (var workbook = new XLWorkbook(Path.Combine(filepath)))
            {
                var worksheet = workbook.Worksheet(1); // Get first worksheet
                var rows = worksheet.RowsUsed();

                foreach (var row in rows)
                {
                    in_count_China++;

                    /* Skip header row (row 1 and 2) */
                    if (row.RowNumber() == 1 || row.RowNumber() == 2)
                        continue;

                    /* SerialNumber cannot be empty */
                    string SerialNumber = row.Cell(1).GetValue<string>();
                    if (String.IsNullOrEmpty(SerialNumber) || SerialNumber == "N/A" || SerialNumber == "NA")
                        continue;

                    Printer printer = new Printer()
                    {
                        Country = "JAP",
                        Class = "Printer",
                        AssetTag = row.Cell(1).GetValue<string>(),
                        SerialNumber = row.Cell(4).GetValue<string>(),
                        AssetStatus = "Active",
                        Location = row.Cell(11).GetValue<string>(),
                        LocationDetail = row.Cell(11).GetValue<string>(),
                        OwnedBy = row.Cell(7).GetValue<string>(),
                        Model = row.Cell(3).GetValue<string>(),
                        SupportGroup = ""
                    };

                    printer_china.Add(printer);
                }
            }

            return printer_china;
        }
    }
}
