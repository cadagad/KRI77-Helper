using KRI77_Helper;
using KRI77_Helper.Utils;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Configuration.Json;
using System.Buffers;
using System.IO;
using static KRI77_Helper.Utils.CsvUtils;
using static System.Net.WebRequestMethods;

#region Configuration Mapping
var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

string? in_path = configuration["FileSettings:In_Path"];
string? out_path = configuration["FileSettings:Out_Path"];
string archive_path = configuration["FileSettings:Archive_Path"] ?? "";

string in_servers = configuration["FileSettings:In_TaniumServers"] ?? "";
string out_servers = configuration["FileSettings:Out_TaniumServers"] ?? "";

string in_eud = configuration["FileSettings:In_TaniumEUD"] ?? "";
string out_eud = configuration["FileSettings:Out_TaniumEUD"] ?? "";

string in_intune = configuration["FileSettings:In_IntuneReport"] ?? "";
string out_intune = configuration["FileSettings:Out_IntuneReport"] ?? "";

string in_terminals = configuration["FileSettings:In_Terminals"] ?? "";
string out_terminals = configuration["FileSettings:Out_Terminals"] ?? "";

string in_network_na = configuration["FileSettings:In_Network_NA"] ?? "";
string in_network_asia = configuration["FileSettings:In_Network_Asia"] ?? "";
string out_network = configuration["FileSettings:Out_Network"] ?? "";

string in_printer_na = configuration["FileSettings:In_Printer_NA"] ?? "";
string in_printer_asia = configuration["FileSettings:In_Network_Asia"] ?? "";
string out_printer = configuration["FileSettings:Out_Printer_ALL"] ?? "";


if (String.IsNullOrEmpty(in_path) || String.IsNullOrEmpty(out_path))
{
    //Console.WriteLine("FileSettings in appsettings.json is not configured properly.");
    CsvLogger.Log("FileSettings in appsettings.json is not configured properly.", level: "ERROR");
    return;
}

if (Directory.Exists(in_path) == false)
{
    //Console.WriteLine($"Input path '{in_path}' does not exist. Please check the configuration.");
    CsvLogger.Log($"Input path '{in_path}' does not exist. Please check the configuration.", level: "ERROR");
    return;
}

if (Directory.Exists(out_path) == false)
{
    Directory.CreateDirectory(out_path);
}
#endregion

/* Get all files from path */
IEnumerable<string> matchingFiles =
        Directory.EnumerateFiles(in_path, "*", SearchOption.TopDirectoryOnly)
            .Where(f => f.EndsWith(".csv") || f.EndsWith(".xlsx"));

Process process = new Process();

string file_network_na = String.Empty;
string file_network_asia = String.Empty;
bool is_network_processed = false;

string file_printer_na = String.Empty;
string file_printer_asia = String.Empty;

bool is_printer_processed = false;
bool is_processed = false;


Console.WriteLine("Process started. (DO NOT CLOSE!!!)");
//CsvLogger.Log("Start KRI77 Helper", level: "START");

if (matchingFiles.Count() == 0)
{
    Console.WriteLine("No files found in the input directory.");
    return;
}

foreach (string file in matchingFiles)
{
    string in_file = Path.GetFileName(file);
    try
    {
        /* Servers - TaniumServers */
        if (in_file.StartsWith(in_servers) && !String.IsNullOrEmpty(in_servers))
        {
            process.ProcessTaniumServers(in_path, in_file, out_path, out_servers, archive_path);
            is_processed = true;
        }

        /* End User Devices - TaniumEUD */
        else if (in_file.StartsWith(in_eud) && !String.IsNullOrEmpty(out_eud))
        {
            process.ProcessTaniumEUD(in_path, in_file, out_path, out_eud, archive_path);
            is_processed = true;
        }

        /* Mobile devices - IntuneReport */
        else if (in_file.StartsWith(in_intune) && !String.IsNullOrEmpty(out_intune))
        {
            process.ProcessIntuneReport(in_path, in_file, out_path, out_intune, archive_path);
            is_processed = true;
        }

        /* Terminals */
        //else if (in_file.StartsWith(in_terminals) && !String.IsNullOrEmpty(out_terminals) && in_file.ToLower().EndsWith(".xlsx"))
        //    process.ProcessTerminals(in_path, in_file, out_path, out_terminals, archive_path);

        /* Network Devices - Process only if both NA and Asia files are present */
        else if (in_file.Contains(in_network_na) && !String.IsNullOrEmpty(out_network))
        {
            file_network_na = in_file;
            is_processed = true;
        }

        else if (in_file.Contains(in_network_asia) && !String.IsNullOrEmpty(out_network))
        {
            file_network_asia = in_file;
            is_processed = true;
        }

        /* Printers */
        //else if (in_file.StartsWith(in_printer_na) && !String.IsNullOrEmpty(out_printer) && in_file.ToLower().EndsWith(".xlsx"))
        //    file_printer_na = in_file;

        /* Process Network Devices if both files are found */
        if (!String.IsNullOrEmpty(file_network_na) && !String.IsNullOrEmpty(file_network_asia) && is_network_processed == false)
        {
            //in_file = file_network_na + ", " + file_network_asia;
            process.ProcessNetworkDevices(in_path, file_network_na, file_network_asia, out_path, out_network, archive_path);
            is_network_processed = true;
            
        }

        /* Process Printer Devices if both files are found */
        //if (!String.IsNullOrEmpty(file_network_na) && !String.IsNullOrEmpty(file_network_asia) && is_network_processed == false)
        //if (!String.IsNullOrEmpty(file_printer_na) && is_printer_processed == false)
        //{
        //    process.ProcessPrinters(archive_path, in_path, file_printer_na, "", out_path, out_printer);
        //    is_printer_processed = true;
        //}

        if (!is_processed) {
            throw new InvalidOperationException($"Invalid Filename - {in_file}");
        }

        is_processed = false;
    }
    catch (Exception ex)
    {
        EmailUtils.SendErrorEmail(in_file, ex.Message, true);
        //Console.Error.WriteLine($"Error: {ex.Message}");
    }
}

if ((!String.IsNullOrEmpty(file_network_na) && String.IsNullOrEmpty(file_network_asia)) || (String.IsNullOrEmpty(file_network_na) && !String.IsNullOrEmpty(file_network_asia)))
{
    //Console.Error.WriteLine($"Error: Missing Network File");
    EmailUtils.SendErrorEmail("Network Devices", "Missing Network File", true);
}

/* Cleanup all files after processing */
Console.WriteLine("\nCleaning up input directory...");
foreach (string file in matchingFiles)
{
    /* Delete file */
    Console.WriteLine($"Deleting file: {file}");
    System.IO.File.Delete(file);
}

//Console.WriteLine($"See output files in the output directory : \n{out_path}");
//CsvLogger.Log($"See output files in the output directory :{out_path}");
//CsvLogger.Log("********************************************************************");

//Console.WriteLine("Processing complete. Press any key to exit.");
//Console.ReadKey();



