using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace KRI77_Helper.Models
{
    public class Server
    {
        public string? ChassisType { get; set; }
        public string? ComputerId { get; set; }
        public required string ComputerName { get; set; }
        public string? SerialNumber { get; set; }
        public string? OsPlatform { get; set; }
        public string? OperatingSystem { get; set; }
        public string? ServicePack { get; set; }
        public string? Manufacturer { get; set; }
        public string? IPAddress { get; set; } 
        public string? CreatedDate { get; set; } 
        public string? UpdatedDate { get; set; }
        public string? LastSeen { get; set; } 
        public string? IsVirtual { get; set; } 
        public string? OSVersion { get; set; } 
        public string? SourceID { get; set; } 
        public string? Model { get; set; }   
        public string? Count { get; set; }

    }
}
