using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Spreadsheet;
using KRI77_Helper.Models;
using KRI77_Helper.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KRI77_Helper.Utils.CsvUtils;

namespace KRI77_Helper
{
    internal class Process
    {

        // Instance-level "global" variable
        private readonly Stopwatch _stopwatch = new Stopwatch();

        public Process() { }

        /* Returns the output file name created */
        public string CopyToArchive(string in_path, string archivePath, string in_file)
        {
            var archived = new List<string>();

            string sourceFile = Path.Combine(in_path, in_file);
            if (!File.Exists(sourceFile))
                throw new FileNotFoundException($"Source file not found: {sourceFile}", sourceFile);


            // Ensure archive folder exists
            Directory.CreateDirectory(archivePath);


            // Determine destination file path
            string destinationFile = Path.Combine(archivePath, in_file);

            // Create a unique, timestamped filename: <name>_yyyyMMdd_HHmmss<ext>
            string name = Path.GetFileNameWithoutExtension(in_file);
            string ext = Path.GetExtension(in_file);
            string stamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            destinationFile = Path.Combine(archivePath, $"{name}_{stamp}{ext}");

            // Copy file (will throw if exists and overwrite==false and no timestamp applied)
            File.Copy(sourceFile, destinationFile);

            return Path.GetFileName(destinationFile);
        }
        /* Processing for Tanium Servers */
        public void ProcessTaniumServers(string in_path, string in_file, string out_path, string out_file, string archive_path)
        {
            _stopwatch.Restart();

            //Console.WriteLine($"Processing file: {in_file}");
            Console.WriteLine($"Processing file: {in_file}");

            if(!in_file.ToLower().EndsWith(".csv"))
            {
                throw new InvalidOperationException("Invalid file type. Expected CSV (.csv)");
            }

            // Archive source file before processing
            string archive_file = CopyToArchive(in_path, archive_path, in_file);

            int in_count = 0;
            int out_count = 0;

            List<Server> servers = new List<Server>();
            List<Server> output = new List<Server>();

            /* Read Tanium Server CSV file and store to List <Server> */
            using (var reader = new StreamReader(Path.Combine(in_path, in_file)))
            {
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    in_count++;

                    if (String.IsNullOrEmpty(line))
                        continue;

                    //string[] values = line.Replace("\"", "").Split(',');
                    string[] values = CsvUtils.SplitCsvLine(line);

                    if (values.Length < 17)
                        continue;

                    Server server = new Server()
                    {
                        ChassisType = values[0],
                        ComputerId = values[1],
                        ComputerName = values[2],
                        SerialNumber = values[3],
                        OsPlatform = values[4],
                        OperatingSystem = values[5],
                        ServicePack = values[6],
                        Manufacturer = values[7],
                        IPAddress = values[8],
                        CreatedDate = values[9],
                        UpdatedDate = values[10],
                        LastSeen = values[11],
                        IsVirtual = values[12],
                        OSVersion = values[13],
                        SourceID = values[14],
                        Model = values[15],
                        Count = values[16]
                    };

                    /* Remove 'Z-VRA-' from ComputerName */
                    if (server.ComputerName.Contains("Z-VRA-"))
                        server.ComputerName = server.ComputerName.Replace("Z-VRA-", "");

                    /* Remove domain from ComputerName */
                    if (server.ComputerName.Contains("."))
                        server.ComputerName = server.ComputerName.Split('.')[0];

                    /* Remove email format from ComputerName */
                    if (server.ComputerName.Contains("@"))
                        server.ComputerName = server.ComputerName.Split('@')[0];

                    servers.Add(server);
                }
            }

            /* Get unique server list */
            var distinctServers =
                (from s in servers.Skip(1) select s.ComputerName).Distinct().ToList();

            /* Join unique server list with original list - select first match only */
            var result = (from ds in distinctServers
                          join sv in servers on ds equals sv.ComputerName into matchedGroup
                          let firstMatch = matchedGroup.FirstOrDefault()
                          select new Server
                          {
                              ChassisType = firstMatch != null ? firstMatch.ChassisType : "",
                              ComputerId = firstMatch != null ? firstMatch.ComputerId : "",
                              ComputerName = ds,
                              SerialNumber = firstMatch != null ? firstMatch.SerialNumber : "",
                              OsPlatform = firstMatch != null ? firstMatch.OsPlatform : "",
                              OperatingSystem = firstMatch != null ? firstMatch.OperatingSystem : "",
                              ServicePack = firstMatch != null ? firstMatch.ServicePack : "",
                              Manufacturer = firstMatch != null ? firstMatch.Manufacturer : "",
                              IPAddress = firstMatch != null ? firstMatch.IPAddress : "",
                              CreatedDate = firstMatch != null ? firstMatch.CreatedDate : "",
                              UpdatedDate = firstMatch != null ? firstMatch.UpdatedDate : "",
                              LastSeen = firstMatch != null ? firstMatch.LastSeen : "",
                              IsVirtual = firstMatch != null ? firstMatch.IsVirtual : "",
                              OSVersion = firstMatch != null ? firstMatch.OSVersion : "",
                              SourceID = firstMatch != null ? firstMatch.SourceID : "",
                              Model = firstMatch != null ? firstMatch.Model : "",
                              Count = firstMatch != null ? firstMatch.Count : ""
                          });

            /* Convert to list */
            output = result.ToList();

            /* Output to csv */
            using (var writer = new StreamWriter(Path.Combine(out_path, out_file)))
            {
                /* write header */
                writer.WriteLine("" +
                    "\"Chassis Type\"," +
                    "\"Computer Id\"," +
                    "\"Computer Name\"," +
                    "\"Serial Number\"," +
                    "\"OS Platform\"," +
                    "\"Operating System\"," +
                    "\"Service Pack\"," +
                    "\"Manufacturer\"," +
                    "\"IP Address\"," +
                    "\"Created Date\"," +
                    "\"Updated Date\"," +
                    "\"Last Seen\"," +
                    "\"Is Virtual\"," +
                    "\"OS Version\"," +
                    "\"Source ID\"," +
                    "\"Model\"," +
                    "\"Count\"");
                foreach (var srv in output)
                {
                    out_count++;
                    writer.WriteLine($"" +
                        $"\"{srv.ChassisType}\"," +
                        $"\"{srv.ComputerId}\"," +
                        $"\"{srv.ComputerName}\"," +
                        $"\"{srv.SerialNumber}\"," +
                        $"\"{srv.OsPlatform}\"," +
                        $"\"{srv.OperatingSystem}\"," +
                        $"\"{srv.ServicePack}\"," +
                        $"\"{srv.Manufacturer}\"," +
                        $"\"{srv.IPAddress}\"," +
                        $"\"{srv.CreatedDate}\"," +
                        $"\"{srv.UpdatedDate}\"," +
                        $"\"{srv.LastSeen}\"," +
                        $"\"{srv.IsVirtual}\"," +
                        $"\"{srv.OSVersion}\"," +
                        $"\"{srv.SourceID}\"," +
                        $"\"{srv.Model}\"," +
                        $"\"{srv.Count}\"");
                }
            }

            _stopwatch.Stop();
            //Console.WriteLine($"DoWork runtime: {_stopwatch.Elapsed}");
            string elapsedCustom = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.fff");

            Console.WriteLine($"Row count {in_file} : {in_count.ToString()}");
            Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            Console.WriteLine($"Duplicates Removed : {(in_count - out_count).ToString()}");
            Console.WriteLine($"Process Runtime: {_stopwatch.Elapsed}");
            Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            Console.WriteLine($"Successfully processed file: {in_file}\n" +
                $"**************************\n\n");


            EmailUtils.SendEmail("Servers", in_count, (in_count - out_count), elapsedCustom, false);
            CsvLogger.Log(in_file, out_file, archive_file, in_count, (in_count - out_count));

            //CsvLogger.Log($"Row count {in_file} : {in_count.ToString()}");
            //CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            //CsvLogger.Log($"Duplicates Removed : {(in_count - out_count).ToString()}");
            //CsvLogger.Log($"Output file : {out_path}{out_file}");
            //CsvLogger.Log($"Successfully processed file: {in_file}");
            //CsvLogger.Log($"*******************************", level: "END");
        }

        /* Processing for Tanium End User Devices */
        public void ProcessTaniumEUD(string in_path, string in_file, string out_path, string out_file, string archive_path)
        {
            _stopwatch.Restart();
            

            //Console.WriteLine($"Processing file: {in_file}");
            Console.WriteLine($"Processing file: {in_file}");

            if (!in_file.ToLower().EndsWith(".csv"))
            {
                throw new InvalidOperationException("Invalid file type. Expected CSV (.csv)");
            }

            // Archive source file before processing
            string archive_file = CopyToArchive(in_path, archive_path, in_file);

            int in_count = 0;
            int out_count = 0;

            List<EndUserDevice> endUserDevices = new List<EndUserDevice>();
            List<EndUserDevice> output = new List<EndUserDevice>();

            /* Read Tanium End User Devices CSV file and store to List <EndUserDevice> */
            using (var reader = new StreamReader(Path.Combine(in_path, in_file)))
            {
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    in_count++;

                    if (String.IsNullOrEmpty(line))
                        continue;

                    //string[] values = line.Replace("\"", "").Split(',');
                    string[] values = CsvUtils.SplitCsvLine(line);

                    if (values.Length < 18)
                        continue;

                    EndUserDevice endUserDevice = new EndUserDevice()
                    {
                        ChassisType = values[0],
                        CpuName = values[1],
                        ComputerId = values[2],
                        ComputerName = values[3],
                        SerialNumber = values[4],
                        OsPlatform = values[5],
                        OperatingSystem = values[6],
                        ServicePack = values[7],
                        Manufacturer = values[8],
                        Model = values[9],
                        IPAddress = values[10],
                        UserName = values[11],
                        CreatedDate = values[12],
                        UpdatedDate = values[13],
                        LastSeen = values[14],
                        IsVirtual = values[15],
                        OSVersion = values[16],
                        Count = values[17]
                    };

                    endUserDevices.Add(endUserDevice);
                }
            }

            /* Get unique server list */
            var distinctEud =
                (from eud in endUserDevices.Skip(1) select eud.SerialNumber).Distinct().ToList();

            /* Join unique server list with original list - select first match only */
            var result = (from deud in distinctEud
                          join eud in endUserDevices on deud equals eud.SerialNumber into matchedGroup
                          let firstMatch = matchedGroup.FirstOrDefault()
                          select new EndUserDevice
                          {
                              ChassisType = firstMatch != null ? firstMatch.ChassisType : "",
                              CpuName = firstMatch != null ? firstMatch.CpuName : "",
                              ComputerId = firstMatch != null ? firstMatch.ComputerId : "",
                              ComputerName = firstMatch != null ? firstMatch.ComputerName : "",
                              SerialNumber = deud,
                              OsPlatform = firstMatch != null ? firstMatch.OsPlatform : "",
                              OperatingSystem = firstMatch != null ? firstMatch.OperatingSystem : "",
                              ServicePack = firstMatch != null ? firstMatch.ServicePack : "",
                              Manufacturer = firstMatch != null ? firstMatch.Manufacturer : "",
                              Model = firstMatch != null ? firstMatch.Model : "",
                              IPAddress = firstMatch != null ? firstMatch.IPAddress : "",
                              UserName = firstMatch != null ? firstMatch.UserName : "",
                              CreatedDate = firstMatch != null ? firstMatch.CreatedDate : "",
                              UpdatedDate = firstMatch != null ? firstMatch.UpdatedDate : "",
                              LastSeen = firstMatch != null ? firstMatch.LastSeen : "",
                              IsVirtual = firstMatch != null ? firstMatch.IsVirtual : "",
                              OSVersion = firstMatch != null ? firstMatch.OSVersion : "",
                              Count = firstMatch != null ? firstMatch.Count : ""
                          });

            /* Convert to list */
            output = result.ToList();

            /* Output to Excel (xlsx) */
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Tanium_EUD");

                /* Write header */
                worksheet.Cell(1, 1).Value = "Chassis Type";
                worksheet.Cell(1, 2).Value = "CPU Name";
                worksheet.Cell(1, 3).Value = "Computer Id";
                worksheet.Cell(1, 4).Value = "Computer Name";
                worksheet.Cell(1, 5).Value = "Serial Number";
                worksheet.Cell(1, 6).Value = "OS Platform";
                worksheet.Cell(1, 7).Value = "Operating System";
                worksheet.Cell(1, 8).Value = "Service Pack";
                worksheet.Cell(1, 9).Value = "Manufacturer";
                worksheet.Cell(1, 10).Value = "Model";
                worksheet.Cell(1, 11).Value = "IP Address";
                worksheet.Cell(1, 12).Value = "User Name";
                worksheet.Cell(1, 13).Value = "Created Date";
                worksheet.Cell(1, 14).Value = "Updated Date";
                worksheet.Cell(1, 15).Value = "Last Seen";
                worksheet.Cell(1, 16).Value = "Is Virtual";
                worksheet.Cell(1, 17).Value = "OS Version";
                worksheet.Cell(1, 18).Value = "Count";

                /* Format header row */
                var headerRow = worksheet.Range(1, 1, 1, 18);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                /* Write data */
                int row = 2;
                foreach (var eud in output)
                {
                    out_count++;
                    worksheet.Cell(row, 1).Value = eud.ChassisType;
                    worksheet.Cell(row, 2).Value = eud.CpuName;
                    worksheet.Cell(row, 3).Value = eud.ComputerId;
                    worksheet.Cell(row, 4).Value = eud.ComputerName;
                    worksheet.Cell(row, 5).Value = eud.SerialNumber;
                    worksheet.Cell(row, 6).Value = eud.OsPlatform;
                    worksheet.Cell(row, 7).Value = eud.OperatingSystem;
                    worksheet.Cell(row, 8).Value = eud.ServicePack;
                    worksheet.Cell(row, 9).Value = eud.Manufacturer;
                    worksheet.Cell(row, 10).Value = eud.Model;
                    worksheet.Cell(row, 11).Value = eud.IPAddress;
                    worksheet.Cell(row, 12).Value = eud.UserName;
                    worksheet.Cell(row, 13).Value = eud.CreatedDate;
                    worksheet.Cell(row, 14).Value = eud.UpdatedDate;
                    worksheet.Cell(row, 15).Value = eud.LastSeen;
                    worksheet.Cell(row, 16).Value = eud.IsVirtual;
                    worksheet.Cell(row, 17).Value = eud.OSVersion;
                    worksheet.Cell(row, 18).Value = eud.Count;
                    row++;
                }

                /* Auto-fit columns */
                worksheet.Columns().AdjustToContents();

                /* Save workbook */
                workbook.SaveAs(Path.Combine(out_path, out_file));
            }

            _stopwatch.Stop();
            //Console.WriteLine($"DoWork runtime: {_stopwatch.Elapsed}");
            string elapsedCustom = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.fff");

            Console.WriteLine($"Row count {in_file} : {in_count.ToString()}");
            Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            Console.WriteLine($"Duplicates Removed : {(in_count - out_count).ToString()}");
            Console.WriteLine($"Process Runtime: {_stopwatch.Elapsed}");
            Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            Console.WriteLine($"Successfully processed file: {in_file}\n" +
                $"**************************\n\n");

            

            EmailUtils.SendEmail("End User Devices", in_count, (in_count - out_count), elapsedCustom, false);
            CsvLogger.Log(in_file, out_file, archive_file, in_count, (in_count - out_count));

            //CsvLogger.Log($"Row count {in_file} : {in_count.ToString()}");
            //CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            //CsvLogger.Log($"Duplicates Removed : {(in_count - out_count).ToString()}");
            //CsvLogger.Log($"Output file : {out_path}{out_file}");
            //CsvLogger.Log($"Successfully processed file: {in_file}");
            //CsvLogger.Log($"*******************************", level: "END");
        }

        /* Processing for IntuneReport - Mobile Devices */
        public void ProcessIntuneReport(string in_path, string in_file, string out_path, string out_file, string archive_path)
        {
            _stopwatch.Restart();
            

            //Console.WriteLine($"Processing file: {in_file}");
            Console.WriteLine($"Processing file: {in_file}");

            if (!in_file.ToLower().EndsWith(".csv"))
            {
                throw new InvalidOperationException("Invalid file type. Expected CSV (.csv)");
            }

            // Archive source file before processing
            string archive_file = CopyToArchive(in_path, archive_path, in_file);

            int in_count = 0;
            int out_count = 0;

            List<MobileDevice> mobileDevices = new List<MobileDevice>();
            List<MobileDevice> output = new List<MobileDevice>();

            /* Read IntuneReport (Mobile devices) CSV file and store to List <MobileDevice> */
            using (var reader = new StreamReader(Path.Combine(in_path, in_file)))
            {
                while (!reader.EndOfStream)
                {
                    string? line = reader.ReadLine();
                    in_count++;

                    if (String.IsNullOrEmpty(line))
                        continue;

                    //string[] values = line.Replace("\"", "").Split(',');
                    string[] values = CsvUtils.SplitCsvLine(line);

                    if (values.Length < 53)
                        continue;

                    MobileDevice mobileDevice = new MobileDevice()
                    {
                        DeviceID = values[0],
                        Devicename = values[1],
                        Enrollmentdate = values[2],
                        Lastcheckin = values[3],
                        AzureADDeviceID = values[4],
                        OSversion = values[5],
                        AzureADregistered = values[6],
                        EASactivationID = values[7],
                        Serialnumber = values[8],
                        Manufacturer = values[9],
                        Model = values[10],
                        EASactivated = values[11],
                        IMEI = values[12],
                        LastEASsynctime = values[13],
                        EASreason = values[14],
                        EASstatus = values[15],
                        Compliancegraceperiodexpiration = values[16],
                        Securitypatchlevel = values[17],
                        WifiMAC = values[18],
                        MEID = values[19],
                        Subscribercarrier = values[20],
                        Totalstorage = values[21],
                        Freestorage = values[22],
                        Managementname = values[23],
                        Category = values[24],
                        UserId = values[25],
                        PrimaryuserUPN = values[26],
                        Primaryuseremailaddress = values[27],
                        Primaryuserdisplayname = values[28],
                        WiFiIPv4Address = values[29],
                        WiFiSubnetID = values[30],
                        Compliance = values[31],
                        Managedby = values[32],
                        Ownership = values[33],
                        Devicestate = values[34],
                        Intuneregistered = values[35],
                        Supervised = values[36],
                        Encrypted = values[37],
                        OS = values[38],
                        SkuFamily = values[39],
                        JoinType = values[40],
                        Phonenumber = values[41],
                        Jailbroken = values[42],
                        ICCID = values[43],
                        EthernetMAC = values[44],
                        CellularTechnology = values[45],
                        ProcessorArchitecture = values[46],
                        EID = values[47],
                        SystemManagementBIOSVersion = values[48],
                        TPMManufacturerId = values[49],
                        TPMManufacturerVersion = values[50],
                        ProductName = values[51],
                        Managementcertificateexpirationdate = values[52],

                    };

                    mobileDevices.Add(mobileDevice);
                }
            }

            /* Get unique server list */
            var distinctMobileDevice =
                (from md in mobileDevices.Skip(1) select md.Serialnumber).Distinct().ToList();

            /* Join unique server list with original list - select first match only */
            var result = (from dmd in distinctMobileDevice
                          join md in mobileDevices on dmd equals md.Serialnumber into matchedGroup
                          let firstMatch = matchedGroup.FirstOrDefault()
                          select new MobileDevice
                          {
                              DeviceID = firstMatch != null ? firstMatch.DeviceID : "",
                              Devicename = firstMatch != null ? firstMatch.Devicename : "",
                              Enrollmentdate = firstMatch != null ? firstMatch.Enrollmentdate : "",
                              Lastcheckin = firstMatch != null ? firstMatch.Lastcheckin : "",
                              AzureADDeviceID = firstMatch != null ? firstMatch.AzureADDeviceID : "",
                              OSversion = firstMatch != null ? firstMatch.OSversion : "",
                              AzureADregistered = firstMatch != null ? firstMatch.AzureADregistered : "",
                              EASactivationID = firstMatch != null ? firstMatch.EASactivationID : "",
                              Serialnumber = dmd,
                              Manufacturer = firstMatch != null ? firstMatch.Manufacturer : "",
                              Model = firstMatch != null ? firstMatch.Model : "",
                              EASactivated = firstMatch != null ? firstMatch.EASactivated : "",
                              IMEI = firstMatch != null ? firstMatch.IMEI : "",
                              LastEASsynctime = firstMatch != null ? firstMatch.LastEASsynctime : "",
                              EASreason = firstMatch != null ? firstMatch.EASreason : "",
                              EASstatus = firstMatch != null ? firstMatch.EASstatus : "",
                              Compliancegraceperiodexpiration = firstMatch != null ? firstMatch.Compliancegraceperiodexpiration : "",
                              Securitypatchlevel = firstMatch != null ? firstMatch.Securitypatchlevel : "",
                              WifiMAC = firstMatch != null ? firstMatch.WifiMAC : "",
                              MEID = firstMatch != null ? firstMatch.MEID : "",
                              Subscribercarrier = firstMatch != null ? firstMatch.Subscribercarrier : "",
                              Totalstorage = firstMatch != null ? firstMatch.Totalstorage : "",
                              Freestorage = firstMatch != null ? firstMatch.Freestorage : "",
                              Managementname = firstMatch != null ? firstMatch.Managementname : "",
                              Category = firstMatch != null ? firstMatch.Category : "",
                              UserId = firstMatch != null ? firstMatch.UserId : "",
                              PrimaryuserUPN = firstMatch != null ? firstMatch.PrimaryuserUPN : "",
                              Primaryuseremailaddress = firstMatch != null ? firstMatch.Primaryuseremailaddress : "",
                              Primaryuserdisplayname = firstMatch != null ? firstMatch.Primaryuserdisplayname : "",
                              WiFiIPv4Address = firstMatch != null ? firstMatch.WiFiIPv4Address : "",
                              WiFiSubnetID = firstMatch != null ? firstMatch.WiFiSubnetID : "",
                              Compliance = firstMatch != null ? firstMatch.Compliance : "",
                              Managedby = firstMatch != null ? firstMatch.Managedby : "",
                              Ownership = firstMatch != null ? firstMatch.Ownership : "",
                              Devicestate = firstMatch != null ? firstMatch.Devicestate : "",
                              Intuneregistered = firstMatch != null ? firstMatch.Intuneregistered : "",
                              Supervised = firstMatch != null ? firstMatch.Supervised : "",
                              Encrypted = firstMatch != null ? firstMatch.Encrypted : "",
                              OS = firstMatch != null ? firstMatch.OS : "",
                              SkuFamily = firstMatch != null ? firstMatch.SkuFamily : "",
                              JoinType = firstMatch != null ? firstMatch.JoinType : "",
                              Phonenumber = firstMatch != null ? firstMatch.Phonenumber : "",
                              Jailbroken = firstMatch != null ? firstMatch.Jailbroken : "",
                              ICCID = firstMatch != null ? firstMatch.ICCID : "",
                              EthernetMAC = firstMatch != null ? firstMatch.EthernetMAC : "",
                              CellularTechnology = firstMatch != null ? firstMatch.CellularTechnology : "",
                              ProcessorArchitecture = firstMatch != null ? firstMatch.ProcessorArchitecture : "",
                              EID = firstMatch != null ? firstMatch.EID : "",
                              SystemManagementBIOSVersion = firstMatch != null ? firstMatch.SystemManagementBIOSVersion : "",
                              TPMManufacturerId = firstMatch != null ? firstMatch.TPMManufacturerId : "",
                              TPMManufacturerVersion = firstMatch != null ? firstMatch.TPMManufacturerVersion : "",
                              ProductName = firstMatch != null ? firstMatch.ProductName : "",
                              Managementcertificateexpirationdate = firstMatch != null ? firstMatch.Managementcertificateexpirationdate : "",

                          });

            /* Convert to list */
            output = result.ToList();

            /* Output to Excel (xlsx) */
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("iOS");

                /* Write header */
                worksheet.Cell(1, 1).Value = "Device ID";
                worksheet.Cell(1, 2).Value = "Device name";
                worksheet.Cell(1, 3).Value = "Enrollment date";
                worksheet.Cell(1, 4).Value = "Last check-in";
                worksheet.Cell(1, 5).Value = "Azure AD Device ID";
                worksheet.Cell(1, 6).Value = "OS version";
                worksheet.Cell(1, 7).Value = "Azure AD registered";
                worksheet.Cell(1, 8).Value = "EAS activation ID";
                worksheet.Cell(1, 9).Value = "Serial number";
                worksheet.Cell(1, 10).Value = "Manufacturer";
                worksheet.Cell(1, 11).Value = "Model";
                worksheet.Cell(1, 12).Value = "EAS activated";
                worksheet.Cell(1, 13).Value = "IMEI";
                worksheet.Cell(1, 14).Value = "Last EAS sync time";
                worksheet.Cell(1, 15).Value = "EAS reason";
                worksheet.Cell(1, 16).Value = "EAS status";
                worksheet.Cell(1, 17).Value = "Compliance grace period expiration";
                worksheet.Cell(1, 18).Value = "Security patch level";
                worksheet.Cell(1, 19).Value = "Wi-Fi MAC";
                worksheet.Cell(1, 20).Value = "MEID";
                worksheet.Cell(1, 21).Value = "Subscriber carrier";
                worksheet.Cell(1, 22).Value = "Total storage";
                worksheet.Cell(1, 23).Value = "Free storage";
                worksheet.Cell(1, 24).Value = "Management name";
                worksheet.Cell(1, 25).Value = "Category";
                worksheet.Cell(1, 26).Value = "UserId";
                worksheet.Cell(1, 27).Value = "Primary user UPN";
                worksheet.Cell(1, 28).Value = "Primary user email address";
                worksheet.Cell(1, 29).Value = "Primary user display name";
                worksheet.Cell(1, 30).Value = "WiFiIPv4Address";
                worksheet.Cell(1, 31).Value = "WiFiSubnetID";
                worksheet.Cell(1, 32).Value = "Compliance";
                worksheet.Cell(1, 33).Value = "Managed by";
                worksheet.Cell(1, 34).Value = "Ownership";
                worksheet.Cell(1, 35).Value = "Device state";
                worksheet.Cell(1, 36).Value = "Intune registered";
                worksheet.Cell(1, 37).Value = "Supervised";
                worksheet.Cell(1, 38).Value = "Encrypted";
                worksheet.Cell(1, 39).Value = "OS";
                worksheet.Cell(1, 40).Value = "SkuFamily";
                worksheet.Cell(1, 41).Value = "JoinType";
                worksheet.Cell(1, 42).Value = "Phone number";
                worksheet.Cell(1, 43).Value = "Jailbroken";
                worksheet.Cell(1, 44).Value = "ICCID";
                worksheet.Cell(1, 45).Value = "EthernetMAC";
                worksheet.Cell(1, 46).Value = "CellularTechnology";
                worksheet.Cell(1, 47).Value = "ProcessorArchitecture";
                worksheet.Cell(1, 48).Value = "EID";
                worksheet.Cell(1, 49).Value = "SystemManagementBIOSVersion";
                worksheet.Cell(1, 50).Value = "TPMManufacturerId";
                worksheet.Cell(1, 51).Value = "TPMManufacturerVersion";
                worksheet.Cell(1, 52).Value = "ProductName";
                worksheet.Cell(1, 53).Value = "Management certificate expiration date";

                /* Format header row */
                var headerRow = worksheet.Range(1, 1, 1, 53);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                /* Write data */
                int row = 2;
                foreach (var md in output)
                {
                    out_count++;
                    worksheet.Cell(row, 1).Value = md.DeviceID;
                    worksheet.Cell(row, 2).Value = md.Devicename;
                    worksheet.Cell(row, 3).Value = md.Enrollmentdate;
                    worksheet.Cell(row, 4).Value = md.Lastcheckin;
                    worksheet.Cell(row, 5).Value = md.AzureADDeviceID;
                    worksheet.Cell(row, 6).Value = md.OSversion;
                    worksheet.Cell(row, 7).Value = md.AzureADregistered;
                    worksheet.Cell(row, 8).Value = md.EASactivationID;
                    worksheet.Cell(row, 9).Value = md.Serialnumber;
                    worksheet.Cell(row, 10).Value = md.Manufacturer;
                    worksheet.Cell(row, 11).Value = md.Model;
                    worksheet.Cell(row, 12).Value = md.EASactivated;
                    worksheet.Cell(row, 13).Value = md.IMEI;
                    worksheet.Cell(row, 14).Value = md.LastEASsynctime;
                    worksheet.Cell(row, 15).Value = md.EASreason;
                    worksheet.Cell(row, 16).Value = md.EASstatus;
                    worksheet.Cell(row, 17).Value = md.Compliancegraceperiodexpiration;
                    worksheet.Cell(row, 18).Value = md.Securitypatchlevel;
                    worksheet.Cell(row, 19).Value = md.WifiMAC;
                    worksheet.Cell(row, 20).Value = md.MEID;
                    worksheet.Cell(row, 21).Value = md.Subscribercarrier;
                    worksheet.Cell(row, 22).Value = md.Totalstorage;
                    worksheet.Cell(row, 23).Value = md.Freestorage;
                    worksheet.Cell(row, 24).Value = md.Managementname;
                    worksheet.Cell(row, 25).Value = md.Category;
                    worksheet.Cell(row, 26).Value = md.UserId;
                    worksheet.Cell(row, 27).Value = md.PrimaryuserUPN;
                    worksheet.Cell(row, 28).Value = md.Primaryuseremailaddress;
                    worksheet.Cell(row, 29).Value = md.Primaryuserdisplayname;
                    worksheet.Cell(row, 30).Value = md.WiFiIPv4Address;
                    worksheet.Cell(row, 31).Value = md.WiFiSubnetID;
                    worksheet.Cell(row, 32).Value = md.Compliance;
                    worksheet.Cell(row, 33).Value = md.Managedby;
                    worksheet.Cell(row, 34).Value = md.Ownership;
                    worksheet.Cell(row, 35).Value = md.Devicestate;
                    worksheet.Cell(row, 36).Value = md.Intuneregistered;
                    worksheet.Cell(row, 37).Value = md.Supervised;
                    worksheet.Cell(row, 38).Value = md.Encrypted;
                    worksheet.Cell(row, 39).Value = md.OS;
                    worksheet.Cell(row, 40).Value = md.SkuFamily;
                    worksheet.Cell(row, 41).Value = md.JoinType;
                    worksheet.Cell(row, 42).Value = md.Phonenumber;
                    worksheet.Cell(row, 43).Value = md.Jailbroken;
                    worksheet.Cell(row, 44).Value = md.ICCID;
                    worksheet.Cell(row, 45).Value = md.EthernetMAC;
                    worksheet.Cell(row, 46).Value = md.CellularTechnology;
                    worksheet.Cell(row, 47).Value = md.ProcessorArchitecture;
                    worksheet.Cell(row, 48).Value = md.EID;
                    worksheet.Cell(row, 49).Value = md.SystemManagementBIOSVersion;
                    worksheet.Cell(row, 50).Value = md.TPMManufacturerId;
                    worksheet.Cell(row, 51).Value = md.TPMManufacturerVersion;
                    worksheet.Cell(row, 52).Value = md.ProductName;
                    worksheet.Cell(row, 53).Value = md.Managementcertificateexpirationdate;

                    row++;
                }

                /* Auto-fit columns */
                worksheet.Columns().AdjustToContents();

                /* Save workbook */
                workbook.SaveAs(Path.Combine(out_path, out_file));
            }

            _stopwatch.Stop();
            //Console.WriteLine($"DoWork runtime: {_stopwatch.Elapsed}");
            string elapsedCustom = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.fff");

            Console.WriteLine($"Row count {in_file} : {in_count.ToString()}");
            Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            Console.WriteLine($"Duplicates Removed : {(in_count - out_count).ToString()}");
            Console.WriteLine($"Process Runtime: {_stopwatch.Elapsed}");
            Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            Console.WriteLine($"Successfully processed file: {in_file}\n" +
                $"**************************\n\n");

            EmailUtils.SendEmail("Intune Report", in_count, (in_count - out_count), elapsedCustom, false);
            CsvLogger.Log(in_file, out_file, archive_file, in_count, (in_count - out_count));

            //CsvLogger.Log($"Row count {in_file} : {in_count.ToString()}");
            //CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            //CsvLogger.Log($"Duplicates Removed : {(in_count - out_count).ToString()}");
            //CsvLogger.Log($"Output file : {out_path}{out_file}");
            //CsvLogger.Log($"Successfully processed file: {in_file}");
            //CsvLogger.Log($"*******************************", level: "END");
        }

        /* Process Terminals */
        public void ProcessTerminals(string in_path, string in_file, string out_path, string out_file, string archive_path)
        {
            _stopwatch.Restart();
            // Archive source file before processing
            string archive_file = CopyToArchive(in_path, archive_path, in_file);

            //Console.WriteLine($"Processing file: {in_file}");
            Console.WriteLine($"Processing file: {in_file}");

            int in_count = 0;
            int out_count = 0;

            List<Terminal> terminals = new List<Terminal>();
            List<Terminal> output = new List<Terminal>();

            /* Read Terminals Excel file and store to List <Terminal> */
            using (var workbook = new XLWorkbook(Path.Combine(in_path, in_file)))
            {
                var worksheet = workbook.Worksheet(1); // Get first worksheet
                var rows = worksheet.RowsUsed();

                foreach (var row in rows)
                {
                    in_count++;

                    /* Skip header row (row 1) */
                    if (row.RowNumber() == 1)
                        continue;

                    string serial = row.Cell(2).GetValue<string>();
                    if (String.IsNullOrEmpty(serial) || serial == "N/A" || serial == "NA")
                        continue;

                    Terminal terminal = new Terminal()
                    {
                        DeviceName = row.Cell(1).GetValue<string>(),
                        Model = row.Cell(8).GetValue<string>(),
                        OSVersion = row.Cell(9).GetValue<string>(),
                        SerialNumber = serial,
                        Location = row.Cell(10).GetValue<string>(),
                        TimeZone = row.Cell(11).GetValue<string>(),
                        Locale = row.Cell(13).GetValue<string>(),
                        Group = row.Cell(17).GetValue<string>()
                    };

                    terminals.Add(terminal);
                }
            }

            /* Get unique server list */
            var distinctTerminals =
                (from s in terminals select s.SerialNumber).Distinct().ToList();

            /* Join unique server list with original list - select first match only */
            var result = (from dt in distinctTerminals
                          join ter in terminals on dt equals ter.SerialNumber into matchedGroup
                          let firstMatch = matchedGroup.FirstOrDefault()
                          select new Terminal
                          {
                              DeviceName = firstMatch != null ? firstMatch.DeviceName : "",
                              Model = firstMatch != null ? firstMatch.Model : "",
                              OSVersion = firstMatch != null ? firstMatch.OSVersion : "",
                              SerialNumber = dt,
                              Location = firstMatch != null ? firstMatch.Location : "",
                              TimeZone = firstMatch != null ? firstMatch.TimeZone : "",
                              Locale = firstMatch != null ? firstMatch.Locale : "",
                              Group = firstMatch != null ? firstMatch.Group : ""

                          });

            /* Convert to list */
            output = result.ToList();

            /* Output to xlsx */
            /* Output to Excel (xlsx) */
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Terminals");

                /* Write header */
                worksheet.Cell(1, 1).Value = "Device Name";
                worksheet.Cell(1, 2).Value = "Model";
                worksheet.Cell(1, 3).Value = "OS Version";
                worksheet.Cell(1, 4).Value = "Serial Number";
                worksheet.Cell(1, 5).Value = "Location";
                worksheet.Cell(1, 6).Value = "OimeZone";
                worksheet.Cell(1, 7).Value = "Locale";
                worksheet.Cell(1, 8).Value = "Group";

                /* Format header row */
                var headerRow = worksheet.Range(1, 1, 1, 8);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                /* Write data */
                int row = 2;
                foreach (var t in output)
                {
                    out_count++;
                    worksheet.Cell(row, 1).Value = t.DeviceName;
                    worksheet.Cell(row, 2).Value = t.Model;
                    worksheet.Cell(row, 3).Value = t.OSVersion;
                    worksheet.Cell(row, 4).Value = t.SerialNumber;
                    worksheet.Cell(row, 5).Value = t.Location;
                    worksheet.Cell(row, 6).Value = t.TimeZone;
                    worksheet.Cell(row, 7).Value = t.Locale;
                    worksheet.Cell(row, 8).Value = t.Group;
                    row++;
                }

                /* Auto-fit columns */
                worksheet.Columns().AdjustToContents();

                /* Save workbook */
                workbook.SaveAs(Path.Combine(out_path, out_file));
            }

            _stopwatch.Stop();
            string elapsedCustom = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.fff");

            Console.WriteLine($"Row count {in_file} : {in_count.ToString()}");
            Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            Console.WriteLine($"Duplicates Removed : {(in_count - out_count).ToString()}");
            Console.WriteLine($"Process Runtime: {_stopwatch.Elapsed}");
            Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            Console.WriteLine($"Successfully processed file: {in_file}\n" +
                $"**************************\n\n");

            

            EmailUtils.SendEmail(in_file, in_count, (in_count - out_count), elapsedCustom, false);
            CsvLogger.Log(in_file, out_file, archive_file, in_count, (in_count - out_count));

            //CsvLogger.Log($"Row count {in_file} : {in_count.ToString()}");
            //CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            //CsvLogger.Log($"Duplicates Removed : {(in_count - out_count).ToString()}");
            //CsvLogger.Log($"Output file : {out_path}{out_file}");
            //CsvLogger.Log($"Successfully processed file: {in_file}");
            //CsvLogger.Log($"*******************************", level: "END");
        }

        /* Process Network NA */
        public void ProcessNetworkDevices(string in_path,
            string in_file_network_na,
            string in_file_network_asia,
            string out_path, string
            out_file, string archive_path)
        {
            _stopwatch.Restart();

            if (!in_file_network_asia.ToLower().EndsWith(".csv"))
            {
                throw new InvalidOperationException("Invalid file type. Expected CSV (.csv)");
            }

            if (!in_file_network_na.ToLower().EndsWith(".xlsx"))
            {
                throw new InvalidOperationException("Invalid file type. Expected Excel File (.xlsx)");
            }

            int in_count_na = 0;
            int in_count_asia = 0;
            int out_count = 0;

            List<Network> networks_na = new List<Network>();
            List<Network> networks_asia = new List<Network>();
            List<Network> output = new List<Network>();

            // Archive source files before processing - NA
            string archive_file_na = CopyToArchive(in_path, archive_path, in_file_network_na);

            /* Read Network NA Excel file and store to List <Network> - networks_na */
            //Console.WriteLine($"Processing file: {in_file_network_na}");
            Console.WriteLine($"Processing file: {in_file_network_na}");

            using (var workbook = new XLWorkbook(Path.Combine(in_path, in_file_network_na)))
            {
                var worksheet = workbook.Worksheet(1); // Get first worksheet
                var rows = worksheet.RowsUsed();

                foreach (var row in rows)
                {
                    in_count_na++;

                    /* Skip header row (row 1) */
                    if (row.RowNumber() == 1)
                        continue;

                    /* hostname cannot be empty */
                    string hostname = row.Cell(1).GetValue<string>();
                    if (String.IsNullOrEmpty(hostname) || hostname == "N/A" || hostname == "NA")
                        continue;

                    /* ip address cannot be empty */
                    string ipaddress = row.Cell(2).GetValue<string>();
                    if (String.IsNullOrEmpty(ipaddress) || ipaddress == "N/A" || ipaddress == "NA")
                        continue;

                    /* Remove domain from hostname */
                    if (hostname.Contains("."))
                        hostname = hostname.Split('.')[0];

                    /* Remove email format from ComputerName */
                    if (hostname.Contains("@"))
                        hostname = hostname.Split('@')[0];

                    Network network = new Network()
                    {
                        Country = "North America",
                        HostName = hostname,
                        SerialNumber = row.Cell(3).GetValue<string>(),
                        ModelName = row.Cell(8).GetValue<string>(),
                        CountryLocation = row.Cell(12).GetValue<string>()
                    };

                    networks_na.Add(network);
                }
            }

            /* Read Network Asia CSV file and store to List <Network> - networks_asia */
            //Console.WriteLine($"Processing file: {in_file_network_asia}");
            Console.WriteLine($"Processing file: {in_file_network_asia}");

            // Archive source files before processing - NA
            string archive_file_asia = CopyToArchive(in_path, archive_path, in_file_network_asia);

            using (var reader = new StreamReader(Path.Combine(in_path, in_file_network_asia)))
            {
                bool isFirstRow = true;
                while (!reader.EndOfStream)
                {
                    if (isFirstRow)
                    {
                        // Skip header row
                        reader.ReadLine();
                        isFirstRow = false;
                        continue;
                    }

                    string? line = reader.ReadLine();
                    in_count_asia++;

                    if (String.IsNullOrEmpty(line))
                        continue;

                    //string[] values = line.Replace("\"", "").Split(',');
                    string[] values = CsvUtils.SplitCsvLine(line);

                    if (values.Length < 12)
                        continue;

                    /* hostname cannot be empty */
                    string hostname = values[0];
                    if (String.IsNullOrEmpty(hostname) || hostname == "N/A" || hostname == "NA")
                        continue;

                    /* ip address cannot be empty */
                    string ipaddress = values[3];
                    if (String.IsNullOrEmpty(ipaddress) || ipaddress == "N/A" || ipaddress == "NA")
                        continue;

                    /* Remove domain from hostname */
                    if (hostname.Contains("."))
                        hostname = hostname.Split('.')[0];

                    /* Remove email format from ComputerName */
                    if (hostname.Contains("@"))
                        hostname = hostname.Split('@')[0];

                    Network network = new Network()
                    {
                        Country = "Asia",
                        HostName = hostname,
                        SerialNumber = values[7],
                        ModelName = values[10],
                        CountryLocation = values[11]
                    };

                    networks_asia.Add(network);
                }
            }

            /* Get unique network list for NA */
            var distinctNetworksNA =
                (from n in networks_na select n.HostName).Distinct().ToList();

            /* Get unique network list for NA */
            var distinctNetworksAsia =
                (from n in networks_asia select n.HostName).Distinct().ToList();

            /* Join unique network list in NA with original list - select first match only */
            var resultNA = (from dn_na in distinctNetworksNA
                            join net in networks_na on dn_na equals net.HostName into matchedGroup
                            let firstMatch = matchedGroup.FirstOrDefault()
                            select new Network
                            {
                                Country = firstMatch != null ? firstMatch.Country : "",
                                HostName = dn_na,
                                SerialNumber = firstMatch != null ? firstMatch.SerialNumber : "",
                                ModelName = firstMatch != null ? firstMatch.ModelName : "",
                                CountryLocation = firstMatch != null ? firstMatch.CountryLocation : ""
                            });

            /* Join unique network list in Asia with original list - select first match only */
            var resultAsia = (from dn_asia in distinctNetworksAsia
                              join net in networks_asia on dn_asia equals net.HostName into matchedGroup
                              let firstMatch = matchedGroup.FirstOrDefault()
                              select new Network
                              {
                                  Country = firstMatch != null ? firstMatch.Country : "",
                                  HostName = dn_asia,
                                  SerialNumber = firstMatch != null ? firstMatch.SerialNumber : "",
                                  ModelName = firstMatch != null ? firstMatch.ModelName : "",
                                  CountryLocation = firstMatch != null ? firstMatch.CountryLocation : ""
                              });


            /* Convert to list - Union all NA + Asia */
            output = resultNA.Concat(resultAsia).ToList();

            /* Output to xlsx */
            /* Output to Excel (xlsx) */
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Network");

                /* Write header */
                worksheet.Cell(1, 1).Value = "Country";
                worksheet.Cell(1, 2).Value = "Hostname";
                worksheet.Cell(1, 3).Value = "Serial Number";
                worksheet.Cell(1, 4).Value = "Model Name";
                worksheet.Cell(1, 5).Value = "CountryLocation";

                /* Format header row */
                var headerRow = worksheet.Range(1, 1, 1, 5);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                /* Write data */
                int row = 2;
                foreach (var t in output)
                {
                    out_count++;
                    worksheet.Cell(row, 1).Value = t.Country;
                    worksheet.Cell(row, 2).Value = t.HostName;
                    worksheet.Cell(row, 3).Value = t.SerialNumber;
                    worksheet.Cell(row, 4).Value = t.ModelName;
                    worksheet.Cell(row, 5).Value = t.CountryLocation;
                    row++;
                }

                /* Auto-fit columns */
                worksheet.Columns().AdjustToContents();

                /* Save workbook */
                workbook.SaveAs(Path.Combine(out_path, out_file));
            }

            _stopwatch.Stop();
            string elapsedCustom = _stopwatch.Elapsed.ToString(@"hh\:mm\:ss\.fff");

            Console.WriteLine($"Row count {in_file_network_na} : {in_count_na.ToString()}");
            Console.WriteLine($"Row count {in_file_network_asia} : {in_count_asia.ToString()}");
            Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            Console.WriteLine($"Duplicates Removed : {((in_count_na + in_count_asia) - out_count).ToString()}");
            Console.WriteLine($"Process Runtime: {_stopwatch.Elapsed}");
            Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            Console.WriteLine($"Successfully processed file/s: {in_file_network_na} and {in_file_network_asia}\n" +
                $"**************************\n\n");

            EmailUtils.SendEmail("Network Devices", (in_count_na + in_count_asia), ((in_count_na + in_count_asia) - out_count), elapsedCustom, false);
            CsvLogger.Log(in_file_network_na, out_file, archive_file_na, in_count_na, (in_count_na - distinctNetworksNA.Count()));
            CsvLogger.Log(in_file_network_asia, out_file, archive_file_asia, in_count_asia, (in_count_asia - distinctNetworksAsia.Count()));

            //CsvLogger.Log($"Row count {in_file_network_na} : {in_count_na.ToString()}");
            //CsvLogger.Log($"Row count {in_file_network_asia} : {in_count_asia.ToString()}");
            //CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            //CsvLogger.Log($"Duplicates Removed : {((in_count_na + in_count_asia) - out_count).ToString()}");
            //CsvLogger.Log($"Output file : {out_path}{out_file}");
            //CsvLogger.Log($"Successfully processed file/s: {in_file_network_na} and {in_file_network_asia}");
            //CsvLogger.Log($"*******************************", level: "END");
        }

        public void ProcessPrinters(string archive_path, string in_path,
            string in_file_printer_na,
            string in_file_printer_asia,
            string out_path, string
            out_file)
        {
            int in_count_na = 0;
            int in_count_asia = 0;
            int out_count = 0;

            List<Printer> printers_na = new List<Printer>();
            List<Printer> printers_asia = new List<Printer>();
            List<Printer> output = new List<Printer>();

            // Archive source file before processing
            CopyToArchive(in_path, archive_path, in_file_printer_na);

            /* Read Network NA Excel file and store to List <Network> - networks_na */
            //Console.WriteLine($"Processing file: {in_file_printer_na}");
            CsvLogger.Log($"Processing file: {in_file_printer_na}");

            using (var workbook = new XLWorkbook(Path.Combine(in_path, in_file_printer_na)))
            {
                var worksheet = workbook.Worksheet("MASTER LIST"); // Get first worksheet
                var rows = worksheet.RowsUsed();

                foreach (var row in rows)
                {
                    in_count_na++;

                    /* Skip header row (row 1) */
                    if (row.RowNumber() == 1)
                        continue;

                    /* hostname cannot be empty */
                    string SerialNumber = row.Cell(11).GetValue<string>();
                    if (String.IsNullOrEmpty(SerialNumber) || SerialNumber == "N/A" || SerialNumber == "NA")
                        continue;

                    Printer printer = new Printer()
                    {
                        Country = row.Cell(1).GetValue<string>(),
                        Class = "Printer",
                        AssetTag = row.Cell(10).GetValue<string>(),
                        SerialNumber = row.Cell(11).GetValue<string>(),
                        AssetStatus = row.Cell(16).GetValue<string>(),
                        Location = row.Cell(22).GetValue<string>(),
                        LocationDetail = row.Cell(1).GetValue<string>(),
                        OwnedBy = "",
                        Model = row.Cell(14).GetValue<string>(),
                        SupportGroup = ""
                    };

                    printers_na.Add(printer);
                }
            }

            /* Get all printers from Asia*/
            printers_asia.AddRange(PrinterUtils.ConsolidatePrintersAsia(in_path));

            /* Get unique NA printer list */
            var distinctNAPrinters =
                (from s in printers_na select s.SerialNumber).Distinct().ToList();

            /* Get unique Asia printer list */
            var distinctAsiaPrinters =
                (from s in printers_asia select s.SerialNumber).Distinct().ToList();

            /* Join unique network list in NA with original list - select first match only */
            var resultNA = (from dp_na in distinctNAPrinters
                            join print in printers_na on dp_na equals print.SerialNumber into matchedGroup
                            let firstMatch = matchedGroup.FirstOrDefault()
                            select new Printer
                            {
                                Country = firstMatch != null ? firstMatch.Country : "",
                                Class = firstMatch != null ? firstMatch.Class : "",
                                AssetTag = firstMatch != null ? firstMatch.AssetTag : "",
                                SerialNumber = dp_na,
                                AssetStatus = firstMatch != null ? firstMatch.AssetStatus : "",
                                Location = firstMatch != null ? firstMatch.Location : "",
                                LocationDetail = firstMatch != null ? firstMatch.LocationDetail : "",
                                OwnedBy = firstMatch != null ? firstMatch.OwnedBy : "",
                                Model = firstMatch != null ? firstMatch.Model : "",
                                SupportGroup = firstMatch != null ? firstMatch.SupportGroup : "",

                            });

            /* Join unique network list in NA with original list - select first match only */
            var resultAsia = (from dp_asia in distinctAsiaPrinters
                              join print in printers_asia on dp_asia equals print.SerialNumber into matchedGroup
                            let firstMatch = matchedGroup.FirstOrDefault()
                            select new Printer
                            {
                                Country = firstMatch != null ? firstMatch.Country : "",
                                Class = firstMatch != null ? firstMatch.Class : "",
                                AssetTag = firstMatch != null ? firstMatch.AssetTag : "",
                                SerialNumber = dp_asia,
                                AssetStatus = firstMatch != null ? firstMatch.AssetStatus : "",
                                Location = firstMatch != null ? firstMatch.Location : "",
                                LocationDetail = firstMatch != null ? firstMatch.LocationDetail : "",
                                OwnedBy = firstMatch != null ? firstMatch.OwnedBy : "",
                                Model = firstMatch != null ? firstMatch.Model : "",
                                SupportGroup = firstMatch != null ? firstMatch.SupportGroup : "",
                            });

            /* Convert to list */
            //output = resultNA.ToList();

            /* Convert to list - Union all NA + Asia */
            output = resultNA.Concat(resultAsia).ToList();

            /* Output to xlsx */
            /* Output to Excel (xlsx) */
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Printers");

                /* Write header */
                worksheet.Cell(1, 1).Value = "Country";
                worksheet.Cell(1, 2).Value = "Class";
                worksheet.Cell(1, 3).Value = "AssetTag";
                worksheet.Cell(1, 4).Value = "SerialNumber";
                worksheet.Cell(1, 5).Value = "AssetStatus";
                worksheet.Cell(1, 6).Value = "Location";
                worksheet.Cell(1, 7).Value = "LocationDetail";
                worksheet.Cell(1, 8).Value = "OwnedBy";
                worksheet.Cell(1, 9).Value = "Model";
                worksheet.Cell(1, 10).Value = "SupportGroup";

                /* Format header row */
                var headerRow = worksheet.Range(1, 1, 1, 10);
                headerRow.Style.Font.Bold = true;
                headerRow.Style.Fill.BackgroundColor = XLColor.LightGray;

                /* Write data */
                int row = 2;
                foreach (var t in output)
                {
                    out_count++;
                    worksheet.Cell(row, 1).Value = t.Country;
                    worksheet.Cell(row, 2).Value = t.Class;
                    worksheet.Cell(row, 3).Value = t.AssetTag;
                    worksheet.Cell(row, 4).Value = t.SerialNumber;
                    worksheet.Cell(row, 5).Value = t.AssetStatus;
                    worksheet.Cell(row, 6).Value = t.Location;
                    worksheet.Cell(row, 7).Value = t.LocationDetail;
                    worksheet.Cell(row, 8).Value = t.OwnedBy;
                    worksheet.Cell(row, 9).Value = t.Model;
                    worksheet.Cell(row, 10).Value = t.SupportGroup;
                    row++;
                }

                /* Auto-fit columns */
                worksheet.Columns().AdjustToContents();

                /* Save workbook */
                workbook.SaveAs(Path.Combine(out_path, out_file));
            }

            //Console.WriteLine($"Row count {in_file_printer_na} : {in_count_na.ToString()}");
            //Console.WriteLine($"Row count {in_file_printer_asia} : {in_count_asia.ToString()}");
            //Console.WriteLine($"Output row count {out_file} : {out_count.ToString()}");
            //Console.WriteLine($"Duplicates Removed : {((in_count_na + in_count_asia) - out_count).ToString()}");
            //Console.WriteLine($"\nOutput file : {out_path}{out_file}\n");
            //Console.WriteLine($"Successfully processed file/s: {in_file_printer_na} and {in_file_printer_asia}\n" +
            //    $"**************************\n\n");

            CsvLogger.Log($"Row count {in_file_printer_na} : {in_count_na.ToString()}");
            CsvLogger.Log($"Row count {in_file_printer_asia} : {in_count_asia.ToString()}");
            CsvLogger.Log($"Output row count {out_file} : {out_count.ToString()}");
            CsvLogger.Log($"Duplicates Removed : {((in_count_na + in_count_asia) - out_count).ToString()}");
            CsvLogger.Log($"Output file : {out_path}{out_file}");
            CsvLogger.Log($"Successfully processed file/s: {in_file_printer_na} and {in_file_printer_asia}");
            CsvLogger.Log($"*******************************", level: "END");
        }

    }
}

